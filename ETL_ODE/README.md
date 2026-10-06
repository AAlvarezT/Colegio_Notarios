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
7. Verifica que aparezca `LISTO PARA GENERAR`.
8. Pulsa `Generar carga`.
9. Verifica que la diferencia sea `S/ 0.00` y que el estado final sea `CUADRADO`.
10. Abre la carpeta de salida.

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

## Archivos generados
- `CARGA_COBRANZAS_ODE_YYYYMM.xls` — **archivo principal**; Excel 97-2003 real generado con `xlwt`. Úsalo en la macro "Pasar ventas y clientes a formato TXT".
- `CARGA_COBRANZAS_ODE_YYYYMM.xlsx` — copia de respaldo en formato moderno; solo para revisión.
- `REPORTE_VALIDACION_ODE_YYYYMM.xlsx` — resumen del proceso (solo revisión, no cargarlo en la macro).

## Ejecutable
Para generar el ejecutable en una construcción limpia, ejecuta:

```bat
generar_exe.bat
```

El ejecutable se genera en `dist`.

## Notas
- Si ya existe un archivo de salida, se agrega fecha y hora al nombre.
- El log técnico queda en `logs/conversor_ode.log`.
