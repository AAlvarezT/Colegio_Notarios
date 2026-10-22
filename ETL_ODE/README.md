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
| 2 | Mapeo de series (`serie_map` en config) | Coincidencia exacta primero, luego prefijo más largo |

Códigos admitidos: `01` = Factura, `03` = Boleta.

Si ninguna fuente resuelve el tipo de un comprobante, **la generación se bloquea** y el registro aparece en el editor de series. No existe respaldo automático: el campo `tipo_documento` en `config.json` es un valor histórico que se conserva por compatibilidad de lectura, pero ya no se usa para clasificar comprobantes.

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

- El campo `tipo_documento` en `config.json` se sigue leyendo, pero **no se usa para clasificar comprobantes**. Es un valor histórico preservado para no romper archivos existentes.
- Reportes anteriores con una sola serie (por ejemplo `FE01`) deben agregar esa serie al `serie_map` o contar con la columna `tipo_comprobante`. El editor de series en la interfaz facilita este paso.

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
- Las cobranzas se registran con la **fecha real de cobro** tomada de la columna `fecha_ingreso` del reporte ODE.
- Los asientos se agrupan por **fecha de cobro + banco/caja + cuenta contable + moneda**.
- No se mezclan cobros de días distintos en un mismo asiento.
- `OTROS` se registra como cobro en efectivo/Caja en la cuenta `10100001`, sin medio de pago.
- Si el rango seleccionado mezcla más de un mes, la generación se bloquea para procesar cada mes por separado.
- `TIP_CAMBIO` se mantiene según la lógica vigente; STARSOFT puede requerir tipos de cambio registrados para las fechas de asiento.
- Los archivos resultantes se guardan en la carpeta `salida`.
- No se debe cerrar el programa mientras está procesando.
- El archivo original no se modifica.
- Al cambiar el archivo, el periodo, el rango de cobro o el correlativo inicial, el análisis previo se invalida automáticamente.

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

---

## Uso web

La interfaz web permite usar el conversor desde cualquier navegador sin instalar Python.

### Inicio local

```bash
# Configura las variables de entorno (copia .env.example a .env y edita)
$env:SECRET_KEY="clave-secreta-aleatoria"
$env:ODE_USERNAME="sharon"
$env:ODE_PASSWORD="tu-clave"

# Inicia el servidor
uvicorn server:app --reload --port 8000
```

Luego abre `http://localhost:8000` en el navegador.

> **Sin credenciales configuradas** (desarrollo local): el servidor permite acceso directo sin login para facilitar pruebas. No usar en producción.

### Variables de entorno

Ver `.env.example` para referencia:

| Variable | Descripción |
|----------|-------------|
| `SECRET_KEY` | Clave para firmar las cookies de sesión. Debe ser aleatoria y secreta. |
| `ODE_USERNAME` | Nombre de usuario para acceder a la interfaz. |
| `ODE_PASSWORD` | Contraseña de acceso. |

### Flujo de uso para Sharon

1. Abre `http://localhost:8000` (o la URL de Vercel).
2. Ingresa las credenciales configuradas.
3. **Sección "Archivos y periodo"**: selecciona el reporte `.xlsx` de ODE, el año y mes del periodo.
4. Define **Fecha cobro desde/hasta** y el **correlativo inicial** para continuar las cargas semanales.
5. **(Opcional)** En "Configuración" ajusta cuentas contables, bancos y serie_map. Puedes exportar/importar la config como JSON.
6. Haz clic en **"Analizar archivos"**: revisa errores, advertencias, rango/abonos/asientos/correlativos y la vista previa.
7. Si hay series sin resolver: selecciona **01 (Factura)** o **03 (Boleta)** para cada una y haz clic en **"Aplicar mapeo y re-analizar"**.
8. Una vez que el estado muestre **"LISTO PARA GENERAR"**, haz clic en **"Generar carga"**.
9. Descarga los archivos:
   - 🟢 **CARGA_ODE_YYYYMM.xls** — este es el archivo para la macro del estudio.
   - **CARGA_ODE_YYYYMM.xlsx** — copia de respaldo.
   - **REPORTE_ODE_YYYYMM.xlsx** — solo revisión interna, no cargarlo en la macro.

### Exportar / importar configuración

- **Exportar**: haz clic en "Exportar config" para descargar la configuración actual como JSON.
- **Importar**: haz clic en "Importar config", selecciona el JSON. El servidor lo valida antes de aplicarlo.

### Configuración en Vercel

1. Crea un nuevo proyecto en Vercel apuntando al repositorio.
2. En **Root Directory**, configura: `ETL_ODE`.
3. El entrypoint se auto-detecta desde `server.py` (variable `app`), o configúralo en `pyproject.toml`: `[tool.vercel] entrypoint = "server:app"`.
4. En **Environment Variables**, define `SECRET_KEY`, `ODE_USERNAME` y `ODE_PASSWORD`.

### Límites

- Archivos de entrada: hasta **8 MB** por archivo.
- Respuesta de generación: hasta ~40 MB (conservador para Vercel).
- No hay almacenamiento persistente — cada request es independiente.

### Pruebas

```bash
# Sólo tests web
python -m pytest tests/test_web.py -v

# Suite completa
python -m pytest tests/ -q
```

**Resultados con datos sintéticos** (sin archivo real): 20/20 tests web pasan.  
**Suite completa**: 138 passed, 1 skipped.

**Pendiente de validar con archivo real de Sharon**: comportamiento con reportes que mezclen facturas y boletas reales, incluyendo formatos de fecha y RUC no estándar.
