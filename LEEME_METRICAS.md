# Polo Oeste - Métricas v1

## Cómo actualizar

1. Detén la aplicación que esté ejecutándose con `Ctrl + C`.
2. Haz una copia de seguridad de `actividad.db`.
3. Reemplaza los archivos de tu proyecto por esta carpeta, sin borrar `actividad.db`.
4. Ejecuta `python app.py` y abre `http://127.0.0.1:5000`.

La base de datos se actualizará automáticamente al iniciar por primera vez.

## Nueva pantalla Métricas

Desde el menú **Métricas** puedes cargar:

- **Consulta de Picking**: cuenta todos los registros con `Cantidad Prep.` y usa
  `Pickeo` como fecha operativa, sin importar el estado del contenedor.
- **Trace de Stock**: toma `REC522` y `REC510` como Almacenaje, `EXP040`
  como Expedición y `REC501` como Recepción. Las unidades se calculan con el
  valor absoluto de Movimiento. Recepción se muestra únicamente como el total
  del mes seleccionado; no es una tarea ni tiene ranking.

Cada carga registra archivos nuevos, duplicados y omitidos. Si el mismo reporte se
carga dos veces, sus movimientos no vuelven a sumarse. Puedes borrar una carga
desde el historial para revertir exclusivamente sus movimientos.

La empresa se conserva en los datos para una futura división por operativa, pero
todavía no aparece como filtro.

Los botones **Ver detalle** de Picking y Almacenaje muestran una búsqueda vacía
al abrirse. El detalle se consulta sólo después de filtrar por fecha, funcionario
o ambos, para no demorar la carga de la pantalla.

La pantalla Métricas abre el último mes que tenga movimientos. El selector de
mes permite consultar períodos anteriores. En los detalles se puede elegir una
fecha desde, una fecha hasta y/o un funcionario.

Cada detalle muestra un resumen por funcionario con las unidades y cantidad de
registros encontrados. La vista limita la tabla a 5.000 filas para mantener su
velocidad; **Exportar Excel** descarga todos los movimientos del filtro en dos
hojas: Resumen por funcionario y Detalle completo.

La importación Bluetooth ya no se muestra en Actividad. La demostración se apoya
en el módulo Handheld web.
