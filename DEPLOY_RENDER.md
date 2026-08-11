# Despliegue en Render

1. Crea un repositorio privado en GitHub y sube el contenido de esta carpeta.
2. En Render, usa **New > Blueprint** y selecciona el repositorio.
3. Render detectará `render.yaml`, creará el servicio web y PostgreSQL.
4. Al solicitar `ADMIN_INITIAL_PASSWORD`, define una contraseña fuerte.
5. Espera el estado **Live** y abre la URL `onrender.com` mostrada por Render.

La primera ejecución crea automáticamente las tablas. La base PostgreSQL es
persistente durante la prueba; no subas `actividad.db` ni la carpeta `uploads`.

> La instancia gratuita es únicamente para prueba: el servicio se suspende tras
> inactividad y la base gratuita vence según las condiciones de Render.
