GENERADOR DE CARGA MASIVA ND · v4.0
COLEGIO DE NOTARIOS DE LIMA

NOVEDADES v4.0
- El programa ya NO genera CSV de validación.
- El entregable operativo queda centrado en TXT/ZIP, como indica el requerimiento.
- Siempre genera dos TXT de control por archivo:
  * <RUC>_ND_<PERIODO>_<CORRELATIVO>_VALIDOS.txt
  * <RUC>_ND_<PERIODO>_<CORRELATIVO>_OBSERVADOS.txt
- Si existe al menos un registro válido, genera además:
  * <RUC>_ND_<PERIODO>_<CORRELATIVO>.txt
  * <RUC>_ND_<PERIODO>_<CORRELATIVO>.zip
  El ZIP oficial incluye SOLO los registros válidos.
- La interfaz permite previsualizar Válidos / Observados / TXT oficial.
- La interfaz permite corregir y revalidar los observados; al corregirlos se regeneran los TXT y el ZIP.
- En el resumen por archivo se muestran directamente las rutas de TXT válidos, TXT observados y ZIP oficial.
- Los reportes Excel/PDF siguen siendo auxiliares de control y no sustituyen el TXT requerido.

FLUJO RECOMENDADO
1. Abra INICIAR_GENERADOR_ND.bat.
2. Agregue uno o varios Registros de Compras .xlsx.
3. Configure la carpeta de salida y la regla de IGV.
4. Presione "Procesar Registro(s) de Compras".
5. Revise el resumen:
   - ND encontrados
   - Válidos
   - Observados
6. Use "Previsualizar TXT" para revisar cada grupo.
7. Use "Corregir / reenviar observados" si corresponde.
8. Cuando existan registros válidos, el ZIP oficial se genera automáticamente con ellos.

IMPORTANTE SOBRE EL IGV
La regla "Usar IGV registrado" toma las columnas I.G.V. del Registro de Compras.
La regla "Calcular porcentaje sobre Adquisiciones no Gravadas" permite aplicar una tasa configurable.
La regla definitiva debe quedar alineada con el criterio contable confirmado por el Colegio.

GENERAR EXE
Ejecute CREAR_EXE.bat.
El ejecutable se genera en la carpeta dist.
