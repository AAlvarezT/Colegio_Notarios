# Conversor de cobranzas ODE

Aplicación para transformar el reporte de comprobantes cancelados de la Notaría Sofía Ode al formato de carga contable.

## Requisitos
- Python 3.11
- Dependencias de [requirements.txt](requirements.txt) (incluye `xlwt==1.3.0`)

## Inicio rápido
1. Descomprime completamente el ZIP del proyecto.
2. Ejecuta `dist\Conversor_ODE.exe`.
3. Selecciona el reporte de cancelados de ODE.
4. Selecciona el Excel modelo del estudio.
5. Indica el periodo.
6. Pulsa `Analizar archivos`.
7. Revisa la sección **Tipos de comprobante detectados** (ver abajo).
8. Verifica que aparezca `LISTO PARA GENERAR`.
9. Pulsa `Generar carga`.
10. Verifica que la diferencia sea `S/ 0.00` y que el estado final sea `CUADRADO`.
11. Abre la carpeta de salida.

## Soporte para Facturas y Boletas en el mismo reporte

El conversor asigna `TIP_DOC` **por comprobante**, no globalmente.  
Después de analizar el archivo, la sección **Tipos de comprobante detectados** muestra:

- Conteo de Facturas (01), Boletas (03) y series sin resolver.
- Tabla de series únicas con su tipo resuelto (verde = ok, naranja = sin resolver).

### Cómo se identifica el tipo por fila

| Prioridad | Fuente | Condición |
|-----------|--------|-----------|
| 1 | Columna `tipo_comprobante` del reporte | Cuando existe y contiene `01` o `03` |
| 2 | Mapeo de series (`serie_map` en config) | Serie coincide con un prefijo configurado |
| 3 | `tipo_documento` global (respaldo) | Solo cuando `serie_map` está vacío (compat. anterior) |

Códigos admitidos: `01` = Factura, `03` = Boleta.  
Un tipo desconocido o un conflicto entre columna y mapeo **bloquea la generación** con mensaje de fila y comprobante.

### Configurar el mapeo de series

Si el reporte no tiene columna de tipo explícita, configure los prefijos de serie en `config.json`:

```json
"serie_map": {
  "FE": "01",
  "BE": "03",
  "F":  "01",
  "B":  "03"
}
```

O use el editor gráfico en la sección **Tipos de comprobante detectados** → campos naranja → **Guardar mapeo y re-analizar**.

### Compatibilidad con configuraciones anteriores

- Configuraciones existentes con `tipo_documento: "01"` y `serie_map` vacío siguen funcionando sin cambios (todos los comprobantes reciben `01`).
- El campo `tipo_documento` ahora se muestra como **(respaldo global)** en la interfaz.
- Al configurar `serie_map`, el respaldo global se desactiva; todas las series deben estar cubiertas por el mapa o por una columna explícita.

## Uso en la macro contable del estudio
El archivo que se debe seleccionar en la herramienta **"Pasar ventas y clientes a formato TXT"** es:

```
CARGA_COBRANZAS_ODE_YYYYMM.xls
```

Pasos:
1. Abre la aplicación contable del estudio y usa **"Pasar ventas y clientes a formato TXT"**.
2. En el selector de archivo (filtro `Archivos XLS`), selecciona `CARGA_COBRANZAS_ODE_YYYYMM.xls`.
3. La herramienta generará un archivo TXT.
4. Carga ese TXT en el módulo contable correspondiente.

> **No cargues directamente el reporte de validación** (`REPORTE_VALIDACION_ODE_YYYYMM.xlsx`): ese archivo es solo de revisión interna.

## Consideraciones del proceso
- Las series `0001` se excluyen automáticamente.
- Interbank, BBVA y Banco de la Nación se procesan por separado.
- `OTROS` se registra como cobro en efectivo/Caja en la cuenta `10100001`, sin medio de pago.
- Los archivos resultantes se guardan en la carpeta `salida`.
- No se debe cerrar el programa mientras está procesando.
- El archivo original no se modifica.
- Al cambiar el archivo, el periodo o el criterio de fecha, el análisis previo se invalida automáticamente.

## Archivos generados
- `CARGA_COBRANZAS_ODE_YYYYMM.xls` — **archivo principal**; Excel 97-2003 real generado con `xlwt`. Úsalo en la macro "Pasar ventas y clientes a formato TXT".
- `CARGA_COBRANZAS_ODE_YYYYMM.xlsx` — copia de respaldo en formato moderno; solo para revisión.
- `REPORTE_VALIDACION_ODE_YYYYMM.xlsx` — resumen del proceso (solo revisión, no cargarlo en la macro). Incluye la hoja **TipoDoc** con el conteo de facturas/boletas.

## Ejecutable
Para generar el ejecutable en una construcción limpia, ejecuta:

```bat
generar_exe.bat
```

El ejecutable se genera en `dist`.

## Notas
- Si ya existe un archivo de salida, se agrega fecha y hora al nombre.
- El log técnico queda en `logs/conversor_ode.log`.
- Los espacios en blanco observados en la herramienta del estudio no han sido reproducidos con archivos sintéticos; pendiente de validar con un reporte real de Sharon que incluya boletas.

