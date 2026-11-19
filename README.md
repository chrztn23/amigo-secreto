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

## Datos del sorteo

Guarda `participantes.xlsx` en `data/`. El archivo debe contener una hoja llamada **Participantes** con las columnas **Nombre** y **Grupo familiar**. Escribe el mismo código para las personas de la misma familia y un código propio para cada participante sin familiares en el sorteo. Completa todas las filas.

La aplicación guarda los resultados en `asignaciones.json` y genera una copia en `asignaciones.xlsx` dentro de la carpeta de datos. Para cambiar de carpeta, usa `AMIGO_SECRETO_DATA_DIR`. Los archivos de datos y la contraseña local están excluidos del repositorio.

## Alojamiento

Configura `ADMIN_PASSWORD` y una `SECRET_KEY` estable en el servicio. Coloca el Excel en almacenamiento privado y persistente, y apunta `AMIGO_SECRETO_DATA_DIR` a esa carpeta. El servicio necesita permiso de escritura para conservar los resultados.

El plan gratuito de Render usa almacenamiento temporal. Para una actividad real, usa almacenamiento persistente para que los resultados se conserven al reiniciar o volver a desplegar. [Documentación de Render](https://render.com/docs/free#local-files-lost-on-redeploy)
