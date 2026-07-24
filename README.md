# Automatización de reporte de Operaciones

Automatiza la incorporación de las operaciones diarias del Reporte 1 al Reporte 2, conservando el histórico y generando un archivo con la fecha del último movimiento.

## Uso en Windows

1. Instalar Python 3.11 o superior.
2. Abrir PowerShell en esta carpeta.
3. Crear y activar un entorno virtual:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

4. Colocar en esta carpeta el archivo `Reporte 1.xlsx` actualizado y el último `Reporte Operaciones DD-MM-AA.xlsx`.
5. Ejecutar `Abrir automatizacion.bat` o abrir `interfaz_reporte.py`.
6. Ingresar la fecha en formato `DD-MM-AAAA`.

El programa genera `Reporte Operaciones DD-MM-AA.xlsx` y conserva los movimientos anteriores.

## Datos no incluidos

Los archivos Excel operativos están excluidos del repositorio porque pueden contener información confidencial de clientes y operaciones. Deben copiarse manualmente en la carpeta local.

## Configuración

Las equivalencias de empresas, ejecutivos, bancos, series documentarias y reglas fiscales se encuentran en `configuracion_reporte.json`.
