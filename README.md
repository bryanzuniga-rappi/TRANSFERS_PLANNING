# Transfer Planner Web

Interfaz web brutalista para ejecutar el modelo de abasto sin editar código.
El archivo `modelo_abasto.py` contiene el motor validado; `app.py` solamente
recopila parámetros, prepara archivos temporales y presenta los resultados.

## Qué cambia para el usuario

- El CSV diario se arrastra desde la computadora; ya no se busca en `TR_PLANS`.
- `DATA_TRANSFERS` puede descargarse automáticamente desde su Google Sheet o
  cargarse manualmente como XLSX.
- Los warehouses origen, su orden, la fecha y el máximo de tareas se editan en
  controles visibles.
- Los resultados se descargan como ZIP o archivo por archivo.
- Ningún input ni output se guarda en Drive durante la ejecución web.

## Arquitectura recomendada

```text
Navegador
   │  CSV por arrastre (hasta 500 MB)
   ▼
Streamlit en Cloud Run (2 CPU, 4 GiB, concurrencia 1)
   ├── exporta DATA_TRANSFERS como XLSX
   ├── ejecuta modelo_abasto.py sin cambiar sus reglas
   └── entrega Reporte_Planeacion.xlsx + BulkCD_*.csv + ZIP
```

Cloud Run es el destino recomendado para producción. Un CSV de más de 100 MB,
el objeto de carga de Streamlit y las estructuras del motor coexistirán en
memoria; por eso el despliegue usa 4 GiB y concurrencia 1. Streamlit admite
200 MB por archivo de forma predeterminada y este proyecto eleva el límite a
500 MB en `.streamlit/config.toml`.

## Ejecución local

Requiere Python 3.12.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
streamlit run app.py
```

En Windows PowerShell, activa el entorno con:

```powershell
.venv\Scripts\Activate.ps1
```

### Credenciales para DATA_TRANSFERS

Hay tres opciones, en este orden:

1. **Cloud Run:** asigna una service account al servicio y comparte el Google
   Sheet `DATA_TRANSFERS` con el correo de esa cuenta como Viewer.
2. **Local o Community Cloud:** agrega la sección `[gcp_service_account]` en
   `.streamlit/secrets.toml`.
3. **Sin credenciales:** selecciona `Subir XLSX manualmente` en Configuración
   avanzada.

Nunca subas `secrets.toml` a GitHub. El `.gitignore` ya lo excluye.

## Despliegue recomendado: GitHub + Cloud Run

1. Crea un repositorio privado en GitHub y sube el contenido de esta carpeta.
2. Crea una service account para el servicio.
3. Habilita Google Drive API en el proyecto de Google Cloud.
4. Comparte `DATA_TRANSFERS` con el correo de esa service account.
5. Autentica `gcloud`, selecciona tu proyecto y ejecuta:

```bash
chmod +x deploy_cloud_run.sh
SERVICE_ACCOUNT="transfer-planner@TU_PROYECTO.iam.gserviceaccount.com" \
  ./deploy_cloud_run.sh
```

El script despliega con:

- 2 CPU.
- 4 GiB de RAM.
- Una planeación concurrente por instancia.
- Timeout de 60 minutos.
- Escalado a cero cuando no se usa.

`--allow-unauthenticated` hace accesible la URL. Para un equipo interno,
configura `APP_PASSWORD` como variable de entorno o elimina esa opción y usa
autenticación de Google Cloud.

## Alternativa rápida: Streamlit Community Cloud

También puede desplegarse directamente desde GitHub. Es útil para validar la
interfaz, pero no es la recomendación principal para corridas constantes con
CSV mayores a 100 MB porque no permite dimensionar explícitamente la memoria
de la instancia.

## Variables visibles

- Warehouses origen, respetando el orden escrito.
- Máximo de tareas.
- Fecha de ejecución.
- Capacidad default por tienda.
- Volumen default por unidad.
- Cantidad mínima positiva.
- ID/URL de `DATA_TRANSFERS`.

## Archivos generados

- `Reporte_Planeacion_DD-MM-YYYY.xlsx`
- Un `BulkCD_<WAREHOUSE_SOURCE>.csv` por origen con asignaciones.
- `Planeacion_DD-MM-YYYY.zip` con todos los archivos anteriores.

## Privacidad y temporales

Los archivos se guardan en el espacio temporal de la instancia únicamente
durante la sesión. Una ejecución nueva elimina el workspace temporal anterior
de esa sesión. Cloud Run usa almacenamiento efímero: al terminar la instancia,
los archivos desaparecen.
