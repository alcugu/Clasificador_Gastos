import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import unicodedata
import pandas as pd
import win32com.client as win32


# -------------------------------------------------------------
# 1. Ventana de Selección de Archivos en Windows (Explorer)
# -------------------------------------------------------------
def seleccionar_fichero_csv() -> str:
  root = tk.Tk()
  root.withdraw()
  root.attributes("-topmost", True)

  ruta_archivo = filedialog.askopenfilename(
      title="Seleccione el archivo CSV de Otros Gastos a clasificar",
      filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
  )
  root.destroy()
  return ruta_archivo


def seleccionar_ruta_guardado(ruta_csv_origen: str) -> str:
  root = tk.Tk()
  root.withdraw()
  root.attributes("-topmost", True)

  carpeta = os.path.dirname(ruta_csv_origen)
  base_nombre = os.path.splitext(os.path.basename(ruta_csv_origen))[0]
  sugerido = f"{base_nombre}_Informe_Dinamico.xlsx"

  ruta_guardado = filedialog.asksaveasfilename(
      title="Guardar archivo Excel resultante como...",
      initialdir=carpeta,
      initialfile=sugerido,
      defaultextension=".xlsx",
      filetypes=[("Libro de Excel (*.xlsx)", "*.xlsx")],
  )
  root.destroy()
  return ruta_guardado


# -------------------------------------------------------------
# 2. Lógica de Clasificación de la Columna S
# -------------------------------------------------------------
def normalizar_texto(texto) -> str:
  if pd.isna(texto):
    return ""
  texto = str(texto).strip()
  texto = unicodedata.normalize("NFD", texto)
  return "".join(c for c in texto if unicodedata.category(c) != "Mn").upper()


def clasificar_linea(row) -> str:
  prov = normalizar_texto(row.get("Proveedor", ""))
  conc = normalizar_texto(row.get("Concepto", ""))
  val_id = row.get("IdProveedor", "")
  id_prov = ""
  if pd.notna(val_id):
    try:
      id_prov = str(int(float(val_id)))
    except (ValueError, TypeError):
      id_prov = str(val_id).strip()

  # 1. Ávoris Retail División (Id 10880)
  if "AVORIS" in prov or id_prov == "10880":
    match = re.search(r"-\s*([A-Z\s]+):", conc)
    servicio = match.group(1).strip() if match else ""
    if servicio == "HOTEL" or "HOTEL" in conc or "ALOJAMIENTO" in conc:
      return "Alojamiento"
    elif (
        servicio == "RENT A CAR"
        or "RENT A CAR" in conc
        or "ALQUILER COCHE" in conc
    ):
      return "Locomoción"
    elif servicio in [
        "AVION",
        "TREN",
        "BUS",
        "GASTO GESTION",
    ] or any(
        k in conc
        for k in [
            "AVION",
            "TREN",
            "BUS",
            "GASTO GESTION",
            "BILLETE",
            "EMISION",
            "VUELO",
        ]
    ):
      return "Viaje"
    return "Viaje"

  # 2. Proveedores conocidos
  if "ASTURIANA DE TRANSPORTE" in prov or id_prov == "26727":
    return "Locomoción"
  if any(p in prov for p in ["BEGIRIS", "DISFRUTON"]) or id_prov in [
      "40094",
      "40508",
  ]:
    return "Manutención"
  if any(
      p in prov for p in ["MICROSOFT", "NALANDA", "DHL", "SBAL", "AMAZON"]
  ) or id_prov in ["17668", "35691", "11582", "20829", "39471"]:
    return "Otros"

  # 3. Conceptos semánticos
  if any(
      k in conc
      for k in [
          "PEAJE",
          "COMBUSTIBLE",
          "GASOIL",
          "GASOLINA",
          "TAXI",
          "RENT A CAR",
          "ALQUILER COCHE",
          "REPARACION",
          "KILOMETRAJE",
      ]
  ):
    return "Locomoción"
  if any(
      k in conc
      for k in [
          "ALIMENTACION",
          "RESTAURANTE",
          "COMIDA",
          "CENA",
          "ALMUERZO",
          "DESAYUNO",
          "DIETA",
      ]
  ):
    return "Manutención"
  if any(k in conc for k in ["HOTEL", "HOSTAL", "ALOJAMIENTO", "HOSPEDAJE"]):
    return "Alojamiento"
  if any(
      k in conc
      for k in [
          "AVION",
          "VUELO",
          "TREN",
          "RENFE",
          "AVE",
          "BILLETE",
          "EMISION",
          "GASTO GESTION",
      ]
  ):
    return "Viaje"

  return "Otros"


# -------------------------------------------------------------
# 3. Generación de Datos y Creación de Tabla Dinámica Real
# -------------------------------------------------------------
def procesar_fichero_con_pivot_table_real(
    fichero_csv: str, fichero_salida: str
):
  # Lectura del CSV
  try:
    df = pd.read_csv(fichero_csv, sep=";", encoding="utf-8-sig")
  except UnicodeDecodeError:
    df = pd.read_csv(fichero_csv, sep=";", encoding="latin1")

  # Clasificación en columna S ("Tipo imputación")
  df["Tipo imputación"] = df.apply(clasificar_linea, axis=1)

  # Convertir importes a formato numérico real (sustituir coma decimal por punto)
  cols_numericas = [
      "#22",
      "#23",
      "Precio Unitario",
      "Descuento",
      "Cant. servida",
      "Pendiente Devolver",
  ]
  for col in cols_numericas:
    if col in df.columns:
      df[col] = (
          df[col]
          .astype(str)
          .str.replace(".", "", regex=False)
          .str.replace(",", ".", regex=False)
      )
      df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

  # 1. Guardar primero la tabla clasificada (A a S) en la hoja "Gastos_Clasificados"
  ruta_abs_excel = os.path.abspath(fichero_salida)
  with pd.ExcelWriter(
      ruta_abs_excel, engine="openpyxl"
  ) as writer:  # noqa: F841
    df.to_excel(writer, sheet_name="Gastos_Clasificados", index=False)

  # 2. Conectar con Microsoft Excel para crear la Tabla Dinámica Nativa
  excel = win32.gencache.EnsureDispatch("Excel.Application")
  excel.Visible = False
  excel.DisplayAlerts = False

  wb = None
  try:
    wb = excel.Workbooks.Open(ruta_abs_excel)
    ws_datos = wb.Sheets("Gastos_Clasificados")

    # Hoja de Informe (se coloca al principio del libro)
    ws_informe = wb.Sheets.Add(Before=ws_datos)
    ws_informe.Name = "Informe"

    # Rango de origen de datos: columnas A a S completas
    num_filas = ws_datos.UsedRange.Rows.Count
    origen_rango = ws_datos.Range(f"A1:S{num_filas}")

    # Constantes oficiales de Excel
    xlDatabase = 1
    xlPageField = 3  # Filtro de informe
    xlRowField = 1  # Filas
    xlColumnField = 2  # Columnas
    xlSum = -4157  # Sumatorio
    xlTabularRow = 1  # Diseño en formato tabular

    # Crear el PivotCache nativo y la Tabla Dinámica
    pivot_cache = wb.PivotCaches().Create(
        SourceType=xlDatabase, SourceData=origen_rango
    )
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=ws_informe.Range("A3"), TableName="TablaDinamicaGastos"
    )

    # 1. Filtros: campo [CC] (columna B)
    campo_filtro = pivot_table.PivotFields("CC")
    campo_filtro.Orientation = xlPageField
    campo_filtro.Position = 1

    # 2. Filas: campo [Tipo imputación] (columna S)
    campo_fila = pivot_table.PivotFields("Tipo imputación")
    campo_fila.Orientation = xlRowField
    campo_fila.Position = 1

    # 3. Sumatorio de Valores: [Suma de #23] y [Suma de #22]
    # Con formato de celda "Estándar" ('General')
    val_23 = pivot_table.AddDataField(
        pivot_table.PivotFields("#23"), "Suma de #23", xlSum
    )
    val_23.NumberFormat = "General"

    val_22 = pivot_table.AddDataField(
        pivot_table.PivotFields("#22"), "Suma de #22", xlSum
    )
    val_22.NumberFormat = "General"

    # 4. Columnas: Utilizar Sumatorio de Valores (DataPivotField)
    try:
      pivot_table.DataPivotField.Orientation = xlColumnField
    except Exception:
      pass

    # 5. Diseño: Mostrar en formato tabular
    pivot_table.RowAxisLayout(xlTabularRow)

    # Ajustar ancho de columnas de la hoja Informe
    ws_informe.Columns("A:D").AutoFit()

    # Guardar cambios
    wb.Save()

  finally:
    if wb is not None:
      wb.Close(SaveChanges=True)
    excel.Quit()


# -------------------------------------------------------------
# 4. Ejecución del Programa
# -------------------------------------------------------------
if __name__ == "__main__":
  archivo_csv = seleccionar_fichero_csv()

  if not archivo_csv:
    print("Operación cancelada: No se seleccionó ningún archivo.")
    sys.exit()

  archivo_excel = seleccionar_ruta_guardado(archivo_csv)

  if not archivo_excel:
    print("Operación cancelada: No se seleccionó la ruta de destino.")
    sys.exit()

  try:
    procesar_fichero_con_pivot_table_real(archivo_csv, archivo_excel)

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(
        "Proceso Completado con Éxito",
        "Se ha generado el archivo Excel con la Tabla Dinámica nativa en:\n\n"
        f"{archivo_excel}\n\n"
        "Al abrir el archivo en Excel podrás modificar campos, filtros y filas"
        " desde el panel de Tabla Dinámica.",
    )
    root.destroy()
    print(f"Archivo generado correctamente: {archivo_excel}")

  except Exception as e:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showerror(
        "Error", f"Ocurrió un error durante la creación del informe:\n\n{str(e)}"
    )
    root.destroy()
    print(f"Error: {e}")