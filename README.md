# 📊 Automatización y Limpieza de Datos de Ventas con Pandas

Proyecto desarrollado en Python para automatizar la lectura, limpieza, transformación y generación de reportes a partir de archivos Excel.

El objetivo principal es tomar un archivo `.xlsx` con información de ventas, procesar los datos mediante Pandas y generar un nuevo archivo Excel con los datos limpios y un resumen de ventas agrupado por sucursal.

## 🚀 Funcionalidades

- 📂 Lectura de archivos `.xlsx`.
- 🔎 Validación de la existencia del archivo.
- 🧹 Eliminación de registros que no poseen un monto de venta.
- ✏️ Reemplazo de sucursales sin identificar.
- 🧽 Eliminación de espacios innecesarios en los nombres de sucursales.
- 🔤 Normalización de los nombres de las sucursales.
- 📊 Agrupación de ventas por sucursal.
- 💰 Cálculo del total vendido por sucursal.
- 🧾 Cálculo de la cantidad de ventas por sucursal.
- 📁 Exportación de los resultados a un nuevo archivo Excel.
- 📑 Generación de dos hojas:
  - `Ventas_Limpias`
  - `Resumen_Sucursales`

## 🛠️ Tecnologías utilizadas

- Python
- Pandas
- OpenPyXL
- Excel (`.xlsx`)

## 📚 Conceptos practicados

Durante el desarrollo del proyecto se trabajó con:

- `pd.read_excel()`
- `DataFrame`
- `dropna()`
- `fillna()`
- `.str.strip()`
- `.str.title()`
- `groupby()`
- `agg()`
- `reset_index()`
- `ExcelWriter`
- `to_excel()`
- Manejo de excepciones con `try/except`
- Funciones para organizar el código
- Manipulación y limpieza de datos tabulares

## 🔄 Funcionamiento

El programa sigue el siguiente flujo:

```text
Archivo Excel
     ↓
Lectura con Pandas
     ↓
Limpieza de datos
     ↓
Normalización de sucursales
     ↓
Agrupación por sucursal
     ↓
Cálculo de estadísticas
     ↓
Generación del reporte
     ↓
reporte_final_gerencia.xlsx
