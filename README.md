# Amigo Secreto

Sorteo familiar en Flask. Cada participante recibe un destinatario único de otro grupo familiar.

## Ejecución local

Con Python 3.12, instala las dependencias y arranca la aplicación.

```sh
python -m venv venv
```

En Windows, ejecuta estos comandos desde PowerShell.

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements.txt
.\venv\Scripts\python.exe -m flask --app app run
```

Abre [Amigo Secreto](http://127.0.0.1:5000/) o el [panel de administración](http://127.0.0.1:5000/panel). Para habilitar el panel, define la variable de entorno `ADMIN_PASSWORD` o guarda la contraseña en `instance/admin-password.txt` antes de arrancar.

## Docker Compose

Coloca `participantes.xlsx` en `data/` y arranca Docker. Desde la carpeta del proyecto, ejecuta:

```sh
docker compose up -d --build
```

Abre [Amigo Secreto](http://127.0.0.1:5000/). Si el puerto está ocupado, añade `AMIGO_HTTP_PORT=5050` a un archivo `.env` local y abre el puerto 5050.

Define `ADMIN_PASSWORD` en `.env` o usa `instance/admin-password.txt` para acceder al panel. Compose guarda las asignaciones en la carpeta de datos del equipo. Si quieres probar mientras el sorteo real sigue abierto, usa una carpeta con datos de prueba mediante `AMIGO_HOST_DATA_DIR` y un puerto distinto.

## Publicar en un VPS con Docker

Los pasos usan un VPS con Ubuntu, acceso SSH y Caddy instalado en el servidor. Docker Compose publica la aplicación en el puerto local `5000`, y Caddy recibe las conexiones públicas por HTTPS.

### 1. Prepara el VPS

Sustituye `usuario` e `IP_DEL_VPS` por tus datos. Entra por SSH y confirma que Docker Engine y su complemento Compose funcionan. Si te faltan, instala ambos con la guía oficial. [Instalar Docker en Ubuntu](https://docs.docker.com/engine/install/ubuntu/)

```sh
ssh usuario@IP_DEL_VPS
```

En el VPS, ejecuta:

```sh
sudo docker compose version
sudo apt update
sudo apt install -y git curl openssl
git clone https://github.com/chrztn23/amigo-secreto.git "$HOME/amigo-secreto"
cd "$HOME/amigo-secreto"
mkdir -p "$HOME/amigo-secreto-carga"
chmod 700 "$HOME/amigo-secreto-carga"
```

Si ya clonaste el repositorio, entra a esa carpeta y usa `git pull --ff-only` para obtener la versión actual. Ejecuta Git con tu usuario de SSH, que necesita permiso de escritura en la carpeta que elijas para el repositorio.

### 2. Sube el Excel y los resultados que quieras continuar

Desde PowerShell en tu equipo, en la carpeta del proyecto, envía el Excel al VPS.

```powershell
scp .\data\participantes.xlsx usuario@IP_DEL_VPS:amigo-secreto-carga/
```

Si quieres continuar con las asignaciones locales, envía también `asignaciones.json`.

```powershell
scp .\data\asignaciones.json usuario@IP_DEL_VPS:amigo-secreto-carga/
```

De vuelta en el VPS, prepara la carpeta que Compose montará. La aplicación se ejecuta con el usuario `10001`.

```sh
sudo install -d -m 0750 -o 10001 -g 10001 /srv/amigo-secreto/data
sudo install -m 0640 -o 10001 -g 10001 \
  "$HOME/amigo-secreto-carga/participantes.xlsx" /srv/amigo-secreto/data/participantes.xlsx
```

Si enviaste las asignaciones, este comando las importa cuando la carpeta de destino todavía no tiene resultados.

```sh
if ! sudo test -e /srv/amigo-secreto/data/asignaciones.json; then
  sudo install -m 0600 -o 10001 -g 10001 \
    "$HOME/amigo-secreto-carga/asignaciones.json" /srv/amigo-secreto/data/asignaciones.json
fi
```

### 3. Configura la contraseña y arranca el contenedor

En la carpeta clonada del VPS, genera el archivo privado `.env` una sola vez. Guarda la contraseña del panel en tu gestor de contraseñas. Si ya tienes un `.env`, edítalo para usar los mismos nombres de variables y conservar tu clave de sesión.

```sh
(
  set -e
  set -C
  umask 077
  admin_password="$(openssl rand -hex 18)"
  secret_key="$(openssl rand -hex 32)"
  printf '%s\n' \
    'AMIGO_HOST_DATA_DIR=/srv/amigo-secreto/data' \
    "ADMIN_PASSWORD=$admin_password" \
    "SECRET_KEY=$secret_key" > .env
)
grep '^ADMIN_PASSWORD=' .env
sudo docker compose up -d --build
sudo docker compose ps
curl -fsS --retry 5 --retry-all-errors --retry-delay 1 \
  -o /dev/null -w 'HTTP %{http_code}\n' http://127.0.0.1:5000/
```

Si el puerto local `5000` está ocupado, añade `AMIGO_HTTP_PORT=5050` a `.env` y usa `5050` en la comprobación y en Caddy.

### 4. Conecta el dominio y habilita HTTPS

Crea un subdominio en Duck DNS y apunta su registro a la IP pública del VPS. Permite conexiones TCP a los puertos `80` y `443` en el firewall del servidor y en el del proveedor, si existe. [Duck DNS](https://www.duckdns.org/) · [Instalar Caddy en Ubuntu](https://caddyserver.com/docs/install#debian-ubuntu-raspbian-packages)

Con Caddy instalado, agrega el siguiente bloque al archivo `/etc/caddy/Caddyfile` y reemplaza el dominio de ejemplo por el tuyo. Si ya tienes otro proxy inverso, configura ese proxy para enviar el tráfico al mismo puerto local. [Guía de Caddy](https://caddyserver.com/docs/quick-starts/reverse-proxy#https-from-client-to-proxy)

```caddyfile
mi-amigo-secreto.duckdns.org {
    reverse_proxy 127.0.0.1:5000
}
```

```sh
sudo caddy validate --config /etc/caddy/Caddyfile
sudo systemctl reload caddy
```

Abre tu dominio con `https://` y comprueba que aparecen los participantes. En `/panel`, entra con la contraseña de `.env`. Si el sitio responde pero no lista participantes, el panel muestra el motivo. Para ver errores del contenedor, usa `sudo docker compose logs --tail=50 app`.

### 5. Actualiza la aplicación

Desde la carpeta clonada en el VPS, ejecuta:

```sh
git pull --ff-only
sudo docker compose up -d --build
```

Las asignaciones se guardan en `/srv/amigo-secreto/data`. Incluye esa carpeta en tus respaldos junto con el `.env` privado.

## Datos del sorteo

Guarda `participantes.xlsx` en `data/`. El archivo debe contener una hoja llamada **Participantes** con las columnas **Nombre** y **Grupo familiar**. Escribe el mismo código para las personas de la misma familia y un código propio para cada participante sin familiares en el sorteo. Completa todas las filas.

La aplicación guarda los resultados en `asignaciones.json` y genera una copia en `asignaciones.xlsx` dentro de la carpeta de datos. Para cambiar de carpeta, usa `AMIGO_SECRETO_DATA_DIR`. Los archivos de datos y la contraseña local están excluidos del repositorio.

## Alojamiento

Configura `ADMIN_PASSWORD` y una `SECRET_KEY` estable en el servicio. Coloca el Excel en almacenamiento privado y persistente, y apunta `AMIGO_SECRETO_DATA_DIR` a esa carpeta. El servicio necesita permiso de escritura para conservar los resultados.

El plan gratuito de Render usa almacenamiento temporal. Para una actividad real, usa almacenamiento persistente para que los resultados se conserven al reiniciar o volver a desplegar. [Documentación de Render](https://render.com/docs/free#local-files-lost-on-redeploy)
