import json
import os
import random
from datetime import datetime

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, url_for

app = Flask(__name__)
app.secret_key = "clave-super-secreta"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
PARTICIPANTS_FILE = os.path.join(DATA_DIR, "participantes.json")
ASSIGNMENTS_FILE = os.path.join(DATA_DIR, "asignaciones.json")
EXCEL_FILE = os.path.join(DATA_DIR, "asignaciones.xlsx")


def load_participants():
    """Return the list of participant names from participantes.json."""
    with open(PARTICIPANTS_FILE, "r", encoding="utf-8") as file:
        data = json.load(file)
    return data.get("names", [])


def load_assignments():
    """Load assignments ensuring the JSON file exists."""
    if not os.path.exists(ASSIGNMENTS_FILE):
        save_assignments([])
        return []

    with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as file:
        data = json.load(file)
    return data.get("assignments", [])


def save_assignments(assignments_list):
    """Persist the assignments list to asignaciones.json."""
    payload = {"assignments": assignments_list}
    with open(ASSIGNMENTS_FILE, "w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def get_available_names(participants, assignments):
    """Return names that have not been used yet."""
    used_names = {item["name"] for item in assignments}
    return [name for name in participants if name not in used_names]


def get_partner_candidates(selected_name, participants, assignments):
    """Return names that can still be assigned as partners (one-time use)."""
    used_partners = {
        item["partner"]
        for item in assignments
        if item.get("partner") and not item["partner"].startswith("SIN PAREJA")
    }
    return [
        name
        for name in participants
        if name != selected_name and name not in used_partners
    ]


def export_to_excel():
    """Export all assignments to an Excel file using pandas."""
    assignments = load_assignments()
    if not assignments:
        # Create an empty DataFrame with the expected columns so the file always exists.
        df = pd.DataFrame(columns=["FechaHora", "Nombre", "Pareja"])
    else:
        df = pd.DataFrame(
            [
                {
                    "FechaHora": item.get("timestamp"),
                    "Nombre": item.get("name"),
                    "Pareja": item.get("partner"),
                }
                for item in assignments
            ]
        )
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")


@app.route("/", methods=["GET"])
def index():
    participants = load_participants()
    assignments = load_assignments()
    available_names = get_available_names(participants, assignments)
    return render_template("index.html", available_names=available_names)


@app.route("/asignar", methods=["POST"])
def asignar():
    selected_name = request.form.get("name", "").strip()

    if not selected_name:
        flash("Debes seleccionar tu nombre antes de continuar.")
        return redirect(url_for("index"))

    participants = load_participants()
    if selected_name not in participants:
        flash("El nombre seleccionado no es válido.")
        return redirect(url_for("index"))

    assignments = load_assignments()
    roulette_names = get_partner_candidates(selected_name, participants, assignments)
    existing_assignment = next((a for a in assignments if a["name"] == selected_name), None)
    if existing_assignment:
        partner = existing_assignment["partner"]
        return render_template(
            "result.html",
            name=selected_name,
            partner=partner,
            roulette_names=roulette_names,
            participants=participants,
        )

    candidates = roulette_names
    if not candidates:
        partner = "SIN PAREJA (no hay más participantes disponibles)"
    else:
        partner = random.choice(candidates)

    assignment = {
        "name": selected_name,
        "partner": partner,
        "timestamp": datetime.now().replace(microsecond=0).isoformat(),
    }
    assignments.append(assignment)
    save_assignments(assignments)
    export_to_excel()

    return render_template(
        "result.html",
        name=selected_name,
        partner=partner,
        roulette_names=roulette_names,
        participants=participants,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
