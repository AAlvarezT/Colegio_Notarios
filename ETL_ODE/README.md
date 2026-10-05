# Conversor de cobranzas ODE

Aplicación para transformar el reporte de comprobantes cancelados de la Notaría Sofía Ode al formato de carga contable.

## Requisitos
- Python 3.11
- Dependencias de [requirements.txt](requirements.txt)

## Inicio rápido
1. Ejecuta `ejecutar.bat`.
2. Selecciona el reporte de ODE.
3. Elige la carpeta de salida.
4. Haz clic en `Analizar archivos`.
5. Revisa la vista previa y validaciones.
6. Haz clic en `Generar carga`.

## Archivos generados
- `CARGA_COBRANZAS_ODE_YYYYMM.xlsx`
- `REPORTE_VALIDACION_ODE_YYYYMM.xlsx`

## Ejecutable
Ejecuta:

```bat
generar_exe.bat
```

El ejecutable se genera en `dist`.

## Notas
- El archivo original no se modifica.
- Si ya existe un archivo de salida, se agrega fecha y hora al nombre.
- El log técnico queda en `logs/conversor_ode.log`.
