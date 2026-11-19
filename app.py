import json
import os
import random
import secrets
from datetime import datetime
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from unicodedata import normalize
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

import pandas as pd
from filelock import FileLock
from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.exceptions import HTTPException

from draw import DrawConfigurationError, has_recipient, plan_draw

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("AMIGO_SECRETO_DATA_DIR") or BASE_DIR / "data").expanduser().resolve()
DATA_DIR.mkdir(parents=True, exist_ok=True)
PARTICIPANTS_FILE = DATA_DIR / "participantes.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "asignaciones.json"
EXCEL_FILE = DATA_DIR / "asignaciones.xlsx"
DRAW_UNAVAILABLE_MESSAGE = "El sorteo está pendiente de revisión por parte de la organización."


def load_admin_password():
    password = os.environ.get("ADMIN_PASSWORD")
    if password:
        return password
    password_file = Path(app.instance_path, "admin-password.txt")
    try:
        return password_file.read_text(encoding="utf-8-sig").rstrip("\r\n")
    except FileNotFoundError:
        return None


ADMIN_PASSWORD = load_admin_password()


def load_participants():
    """Read the roster and family codes from the saved participants workbook."""
    try:
        rows = pd.read_excel(
            PARTICIPANTS_FILE,
            sheet_name="Participantes",
            dtype=str,
            keep_default_na=False,
            engine="openpyxl",
        )
    except (OSError, ValueError, BadZipFile, ParseError) as error:
        raise DrawConfigurationError(
            "Abre el archivo participantes.xlsx y verifica que tenga una hoja llamada «Participantes»."
        ) from error

    columns = ["Nombre", "Grupo familiar"]
    if not set(columns).issubset(rows.columns):
        raise DrawConfigurationError("Incluye las columnas «Nombre» y «Grupo familiar» en el Excel de participantes.")

    participants = {}
    for name, group in rows[columns].itertuples(index=False, name=None):
        name = name.strip()
        group = " ".join(normalize("NFKC", group).split()).casefold()
        if not name and group:
            raise DrawConfigurationError("Hay un grupo familiar sin nombre en el Excel de participantes.")
        if name in participants:
            raise DrawConfigurationError(f"El nombre {name} está repetido en el Excel de participantes.")
        if name:
            participants[name] = group
    if not participants:
        raise DrawConfigurationError("Agrega los nombres en el Excel de participantes.")
    return participants


def _assignments_lock():
    return FileLock(f"{ASSIGNMENTS_FILE}.lock")


def load_assignments():
    try:
        with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as file:
            return json.load(file).get("assignments", [])
    except FileNotFoundError:
        return []


def _write_atomic(path, content):
    temporary = NamedTemporaryFile(dir=Path(path).parent, delete=False)
    try:
        with temporary:
            temporary.write(content)
        os.replace(temporary.name, path)
    finally:
        Path(temporary.name).unlink(missing_ok=True)


def save_assignments(assignments_list):
    """Commit the draw and refresh the Excel copy."""
    excel = export_to_excel(assignments_list)
    payload = json.dumps({"assignments": assignments_list}, ensure_ascii=False, indent=2)
    _write_atomic(ASSIGNMENTS_FILE, payload.encode("utf-8"))
    try:
        _write_atomic(EXCEL_FILE, excel.getvalue())
    except OSError:
        app.logger.warning("No se pudo actualizar la copia local de asignaciones.xlsx.", exc_info=True)


def get_available_names(participants, assignments):
    used_names = {item["name"] for item in assignments if has_recipient(item)}
    return [name for name in participants if name not in used_names]


def export_to_excel(assignments):
    rows = [(item.get("timestamp"), item.get("name"), item.get("partner")) for item in assignments]
    excel = BytesIO()
    pd.DataFrame(rows, columns=["FechaHora", "Nombre", "Pareja"]).to_excel(
        excel, index=False, engine="openpyxl"
    )
    excel.seek(0)
    return excel


def draw_status(assignments):
    participants = {}
    try:
        participants = load_participants()
        plan_draw(participants, assignments)
    except DrawConfigurationError as error:
        return participants, str(error)
    return participants, None


def _get_json_payload():
    """Return parsed JSON payload or abort if missing."""
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        abort(400, description="Se requiere un objeto JSON.")
    return data


def _password_matches(provided):
    return (
        bool(ADMIN_PASSWORD)
        and isinstance(provided, str)
        and secrets.compare_digest(provided.encode("utf-8"), ADMIN_PASSWORD.encode("utf-8"))
    )


def _require_password(provided):
    """Abort the request if the password is incorrect."""
    if not _password_matches(provided):
        abort(401, description="Contraseña inválida.")


def _admin_csrf_token():
    if "admin_csrf" not in session:
        session["admin_csrf"] = secrets.token_urlsafe(32)
    return session["admin_csrf"]


def _require_admin_csrf():
    supplied = request.form.get("csrf_token", "").encode("utf-8")
    expected = session.get("admin_csrf", "").encode("utf-8")
    if not supplied or not secrets.compare_digest(supplied, expected):
        abort(400, description="La sesión del formulario caducó. Abre el panel e inténtalo de nuevo.")


def _required_text(data, field, message):
    value = data.get(field, "")
    if not isinstance(value, str) or not value.strip():
        abort(400, description=message)
    return value.strip()


def _delete_assignment(name):
    with _assignments_lock():
        assignments = load_assignments()
        remaining = [item for item in assignments if item["name"] != name]
        if len(remaining) == len(assignments):
            abort(404, description="No existe una asignación para ese nombre.")
        save_assignments(remaining)


def _update_assignment(name, partner):
    with _assignments_lock():
        assignments = load_assignments()
        assignment = next((item for item in assignments if item["name"] == name), None)
        if not assignment:
            abort(404, description="No existe una asignación para ese nombre.")
        assignment["partner"] = partner
        assignment["timestamp"] = datetime.now().replace(microsecond=0).isoformat()
        try:
            participants = load_participants()
            if partner not in participants or not has_recipient(assignment):
                abort(400, description="Selecciona como pareja un nombre del Excel de participantes.")
            plan_draw(participants, assignments)
        except DrawConfigurationError as error:
            abort(409, description=str(error))
        save_assignments(assignments)
        return assignment


def _admin_response(response):
    response.headers["Cache-Control"] = "no-store"
    return response


@app.route("/", methods=["GET"])
def index():
    with _assignments_lock():
        assignments = load_assignments()
    participants, issue = draw_status(assignments)
    available_names = get_available_names(participants, assignments)
    random.shuffle(available_names)
    return render_template(
        "index.html",
        available_names=available_names,
        draw_notice=DRAW_UNAVAILABLE_MESSAGE if issue else None,
    )


@app.route("/asignar", methods=["POST"])
def asignar():
    selected_name = request.form.get("name", "").strip()

    if not selected_name:
        flash("Debes seleccionar tu nombre antes de continuar.")
        return redirect(url_for("index"))

    try:
        with _assignments_lock():
            participants = load_participants()
            if selected_name not in participants:
                flash("El nombre seleccionado no es válido.")
                return redirect(url_for("index"))

            assignments = load_assignments()
            existing = next(
                (item for item in assignments if item["name"] == selected_name and has_recipient(item)),
                None,
            )
            if existing:
                partner = existing["partner"]
                family = participants[selected_name]
                if partner not in participants or (family and family == participants[partner]):
                    raise DrawConfigurationError()
            else:
                partner = plan_draw(participants, assignments, shuffle=True)[selected_name]
                # Replace an old «SIN PAREJA» entry only when the new draw succeeds.
                assignments = [item for item in assignments if item["name"] != selected_name]
                assignments.append({
                    "name": selected_name,
                    "partner": partner,
                    "timestamp": datetime.now().replace(microsecond=0).isoformat(),
                })
                save_assignments(assignments)
    except DrawConfigurationError:
        flash(DRAW_UNAVAILABLE_MESSAGE)
        return redirect(url_for("index"))

    return render_template(
        "result.html",
        name=selected_name,
        partner=partner,
        participants=list(participants),
    )


@app.route("/panel", methods=["GET", "POST"])
def admin_panel():
    if request.method == "POST":
        _require_admin_csrf()
        if not ADMIN_PASSWORD:
            flash("El acceso de administración aún no está configurado.", "error")
        elif _password_matches(request.form.get("password")):
            session["admin"] = True
            session.pop("admin_csrf", None)
        else:
            flash("Contraseña incorrecta.", "error")
        return redirect(url_for("admin_panel"), code=303)

    signed_in = session.get("admin", False)
    assignments = []
    draw_issue = None
    if signed_in:
        with _assignments_lock():
            assignments = load_assignments()
        _, draw_issue = draw_status(assignments)
    page = render_template(
        "admin.html",
        signed_in=signed_in,
        assignments=assignments,
        draw_issue=draw_issue,
        csrf_token=_admin_csrf_token(),
    )
    return page, 200, {"Cache-Control": "no-store"}


@app.post("/panel/asignaciones")
def panel_assignments():
    if not session.get("admin"):
        return redirect(url_for("admin_panel"), code=303)
    _require_admin_csrf()
    try:
        name = _required_text(request.form, "name", "Falta el nombre del participante.")
        action = request.form.get("action")
        if action == "delete":
            _delete_assignment(name)
            flash(f"Asignación de {name} eliminada.", "success")
        elif action == "save":
            partner = _required_text(request.form, "partner", "La pareja no puede estar vacía.")
            _update_assignment(name, partner)
            flash(f"Pareja de {name} actualizada.", "success")
        else:
            abort(400, description="Acción inválida.")
    except HTTPException as error:
        flash(error.description, "error")
    return redirect(url_for("admin_panel"), code=303)


@app.post("/panel/salir")
def admin_logout():
    _require_admin_csrf()
    session.pop("admin", None)
    session.pop("admin_csrf", None)
    return redirect(url_for("admin_panel"), code=303)


@app.route("/admin/asignaciones", methods=["GET", "POST", "DELETE"])
def admin_assignments():
    """Simple admin interface to consultar, actualizar o eliminar asignaciones."""
    if request.method == "GET":
        password = request.args.get("password")
        _require_password(password)
        with _assignments_lock():
            assignments = load_assignments()
        return _admin_response(jsonify({"assignments": assignments}))

    data = _get_json_payload()
    _require_password(data.get("password"))
    name = _required_text(data, "name", "El campo 'name' es obligatorio.")
    if request.method == "DELETE":
        _delete_assignment(name)
        return jsonify({"status": "eliminado", "name": name})

    partner = _required_text(data, "partner", "El campo 'partner' es obligatorio para actualizar.")
    assignment = _update_assignment(name, partner)
    return jsonify({"status": "actualizado", "assignment": assignment})


@app.route("/admin/asignaciones/excel", methods=["GET"])
def admin_download_excel():
    """Descargar el archivo Excel con todas las asignaciones."""
    if not session.get("admin"):
        _require_password(request.args.get("password"))
    with _assignments_lock():
        assignments = load_assignments()
    response = send_file(
        export_to_excel(assignments),
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name="asignaciones.xlsx",
    )
    return _admin_response(response)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
