import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Función para leer el archivo txt y procesar las columnas
def procesar_archivo(filepath):
    # Definición de la estructura según la longitud de cada campo
    columnas = [
        ("EMPRESA", 1, 4), ("CLAVE DE INTERFASE", 5, 7), ("FECHA CONTABLE", 8, 15),
        ("FECHA DE OPERACION", 16, 23), ("PRODUCTO", 24, 25), ("SUBPRODUCTO", 26, 29),
        ("GARANTIA", 30, 32), ("TIPO DE PLAZO", 33, 33), ("PLAZO", 34, 36),
        ("SUBSECTOR", 37, 37), ("SECTOR B.E.", 38, 39), ("CNAE", 40, 44),
        ("EMPRESA TUTELADA", 45, 48), ("AMBITO", 49, 50), ("MOROSIDAD", 51, 51),
        ("INVERSION", 52, 52), ("OPERACION", 53, 55), ("CODIGO CONTABLE", 56, 60),
        ("DIVISA", 61, 63), ("TIPO DE DIVISA", 64, 64), ("TIPO NOMINAL", 65, 69),
        ("FILLER", 70, 74), ("VARIOS", 75, 104), ("CLAVE DE AUTORIZACION", 105, 110),
        ("CENTRO OPERANTE", 111, 114), ("CENTRO ORIGEN", 115, 118), ("CENTRO DESTINO", 119, 122),
        ("NUM.MOVTOS AL DEBE", 123, 129), ("NUM.MOVTOS AL HABER", 130, 136),
        ("IMPORTE DEBE EN PESETAS", 137, 151), ("IMPORTE HABER EN PESETAS", 152, 166),
        ("IMPORTE DEBE EN DIVISA", 167, 181), ("IMPORTE HABER EN DIVISA", 182, 196),
        ("INDICADOR DE CORRECCION", 197, 197), ("NUMERO DE CONTROL", 198, 209),
        ("CLAVE DE CONCEPTO", 210, 212), ("DESCRIPCION DE CONCEPTO", 213, 226),
        ("TIPO DE CONCEPTO", 227, 227), ("OBSERVACIONES", 228, 257), ("SANCTCCC", 258, 275),
        ("APLICACION ORIGEN", 276, 278), ("APLICACION DESTINO", 279, 281),
        ("OBSERVACIONES3", 282, 287), ("RESERVAT", 288, 291), ("HACTRGEN", 292, 295),
        ("HAYCOCAI", 296, 296), ("HAYCTORD", 297, 297), ("SATINTER", 298, 302),
        ("SACCLVOP", 303, 305), ("SACCEGES", 306, 309), ("SACAPLCP", 310, 311),
        ("SACCDTGT", 312, 313), ("SAYUTILI", 314, 314), ("SAYROTAC", 315, 316),
        ("FECHA ALTA DE PARTIDA", 317, 324), ("OBSERV4", 325, 354), ("FILLER2", 355, 380)
    ]
    
    data = []
    with open(filepath, 'r') as file:
        for line in file:
            row = {col[0]: line[col[1]-1:col[2]].strip() for col in columnas}
            
            # Agregar el símbolo "$" y formato en las columnas especificadas
            for col_name in ["IMPORTE DEBE EN PESETAS", "IMPORTE HABER EN PESETAS", "IMPORTE DEBE EN DIVISA", "IMPORTE HABER EN DIVISA"]:
                # Formatear los valores numéricos con "$" y dos decimales
                try:
                    # Dividir por 100 y luego aplicar el formato de moneda
                    row[col_name] = f"${float(row[col_name]) / 100:,.2f}"
                except ValueError:
                    # Si el valor no es numérico (por ejemplo, vacío), se deja como está
                    row[col_name] = "$0.00"
                    
            data.append(row)
    
    return pd.DataFrame(data)

# Función para cargar el archivo y procesarlo
def cargar_archivo():
    filepath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if filepath:
        try:
            df = procesar_archivo(filepath)
            exportar_a_excel(df)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {e}")

# Función para exportar los datos procesados a un archivo Excel
def exportar_a_excel(df):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Éxito", "Archivo exportado exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar a Excel: {e}")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Procesador de TXT a Excel")
root.geometry("400x200")

label = tk.Label(root, text="Cargar archivo TXT:")
label.pack(pady=10)

boton_cargar = tk.Button(root, text="Cargar archivo", command=cargar_archivo)
boton_cargar.pack(pady=20)

root.mainloop()
