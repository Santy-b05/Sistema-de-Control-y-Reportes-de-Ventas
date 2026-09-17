import pandas as pd

def AccederArchivo():
    verificacion = False
    while not verificacion:
        try:   
            nombreArchivo = input("Ingrese el nombre del archivo (Debe ser extensión .xlsx): ")
            Excel_dataframe = pd.read_excel(f"{nombreArchivo}.xlsx")
            verificacion = True
            return Excel_dataframe
        except FileNotFoundError:
            print("Error al buscar el archivo, verifique el nombre.")

def ReemplazarNulos(DatosExcel):
    DatosExcel = DatosExcel.dropna(subset=["Monto"])
    DatosExcel["Sucursal"] =  DatosExcel["Sucursal"].fillna(value="Sin Identificar")
    return DatosExcel

def ModificarDatos(DatosExcel):
    DatosExcel["Sucursal"] = DatosExcel["Sucursal"].str.title().str.strip()
    return DatosExcel

def ResumenDeVentas(Datos_Excel):
    TablaResumen = Datos_Excel.groupby("Sucursal")["Monto"].agg(
        Total_Vendido = "sum",
        Cantidad_Ventas = "count"
        ).reset_index()
    return TablaResumen

def Exportacion_Excel(Datos, Ventas):
    nombre_arch = "reporte_final_gerencia.xlsx"
    try:
        with pd.ExcelWriter(nombre_arch, engine="openpyxl") as writer:
            Datos.to_excel(writer, sheet_name="Ventas_Limpias", index=False)
            Ventas.to_excel(writer, sheet_name="Resumen_Sucursales", index=False)
    except PermissionError:
        print("Error de permisos: Verifique de no tener el archivo abierto.")
    except Exception as e:
        print(f"Error al exportar el archivo: {e}")

Datos_Excel = AccederArchivo()
Datos_Excel = ReemplazarNulos(Datos_Excel)
Datos_Excel = ModificarDatos(Datos_Excel)
Ventas = ResumenDeVentas(Datos_Excel)
Exportacion_Excel(Datos_Excel, Ventas)