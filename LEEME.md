# Publicación del dashboard Contribución SENATI

## Contenido de la carpeta

| Archivo | Para qué sirve |
|---|---|
| `dashboard_contribucion.py` | El dashboard |
| `.streamlit/config.toml` | Colores y tema |
| `requirements.txt` | Librerías que necesita |
| `iniciar_dashboard.bat` | Arranque en Windows para compartir en la red |
| `Dockerfile` | Para que TI lo despliegue en un servidor o en Azure |

**Falta agregar:** el archivo de datos `BASE_CONTRI(1).xlsx`, con ese nombre exacto, en esta misma carpeta.

## Opción A — Compartir desde un equipo en la red de SENATI

Se hace una sola vez en el equipo que servirá el dashboard (los usuarios no instalan nada):

1. Copiar esta carpeta completa al equipo y agregar `BASE_CONTRI(1).xlsx`.
2. Instalar las librerías: abrir una terminal en la carpeta y ejecutar `pip install -r requirements.txt`.
3. (Opcional) Abrir `iniciar_dashboard.bat` con el Bloc de notas y escribir una clave en la línea `set DASHBOARD_CLAVE=`. Si se deja vacía, no se pide clave.
4. Hacer doble clic en `iniciar_dashboard.bat`. La ventana muestra las direcciones IP del equipo.
5. Si Windows pregunta por el firewall, permitir el acceso en redes privadas o de dominio.
6. Compartir con los usuarios la dirección `http://IP-DEL-EQUIPO:8501`. Deben estar en la red de SENATI o conectados por VPN.

A tener en cuenta:

- El dashboard funciona mientras el equipo esté encendido y la ventana abierta.
- Si la IP del equipo cambia, la dirección cambia. Se puede pedir a TI una IP fija o un nombre de equipo.
- Si los usuarios no logran entrar, lo habitual es que el puerto 8501 esté bloqueado; TI puede habilitarlo.

## Opción B — Despliegue permanente con TI (servidor o Azure)

Entregar esta carpeta a TI. Datos que necesitan:

- Aplicación Streamlit (Python 3.12), puerto 8501, usa WebSockets.
- Se construye con el `Dockerfile` incluido, o directamente con `pip install -r requirements.txt` y el comando del `Dockerfile`.
- Variable de entorno `DASHBOARD_CLAVE` para la clave de acceso. Recomendado: reemplazarla por autenticación con cuentas institucionales (Microsoft Entra ID) delante de la app.
- Publicar solo con HTTPS y acceso restringido a personal autorizado: el dashboard muestra pagos por contribuyente con RUC.
- Para insertarlo en una página de SharePoint: elemento web "Insertar" con la dirección HTTPS, autorizando el dominio en la configuración del sitio.

## Actualizar los datos

Reemplazar `BASE_CONTRI(1).xlsx` por la versión nueva (mismo nombre y mismas hojas BASE y PRESUPUESTO) y reiniciar el dashboard, o usar "Clear cache" en el menú de la app.
