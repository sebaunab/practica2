import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def consolidar_hojas(excel_path, ruta_salida):
    # Clave de unión fija con conversión a minúsculas
    clave_union = "documento compras"

    hojas_excluidas = ["FACTURAS PROVEEDOR", "Resumen x Mes", "TODOS"]
    excel_data = pd.ExcelFile(excel_path)
    planilla_final = None

    for sheet_name in excel_data.sheet_names:
        if sheet_name in hojas_excluidas:
            print(f"Saltando la hoja: {sheet_name}")
            continue

        planilla = pd.read_excel(excel_path, sheet_name=sheet_name)
        print(f"Hojas: {sheet_name} - Columnas: {planilla.columns.tolist()}")

        # Convertir columnas a minúsculas para uniformidad
        planilla.columns = planilla.columns.str.strip()  # Eliminar espacios
        planilla.columns = planilla.columns.str.replace(r"[\(\)\[\]\.\-]", "", regex=True)  # Eliminar caracteres especiales
        planilla.columns = planilla.columns.str.lower()  # Convertir a minúsculas

        if clave_union not in planilla.columns:
            print(f"Advertencia: La hoja '{sheet_name}' no contiene la columna clave '{clave_union}'.")
            continue

        # Asegurar tipos consistentes para la clave de unión
        planilla[clave_union] = planilla[clave_union].astype(str)
        if planilla_final is not None:
            planilla_final[clave_union] = planilla_final[clave_union].astype(str)

        if planilla_final is None:
            planilla_final = planilla
        else:
            common_columns = planilla_final.columns.intersection(planilla.columns).tolist()
            if clave_union in common_columns:
                common_columns.remove(clave_union)  # Exceptuar la columna clave

            for col in common_columns:
                planilla.rename(columns={col: f"{col}_{sheet_name}"}, inplace=True)

            try:
                planilla_final = pd.merge(planilla_final, planilla, on=clave_union, how='outer')
            except Exception as e:
                print(f"Error al combinar hojas: {e}")
                messagebox.showerror("Error", f"Error al combinar hojas: {e}")
                return

    if planilla_final is not None:
        try:
            planilla_final.to_excel(ruta_salida, index=False)
            print(f"Consolidación completa. Archivo final guardado en: {ruta_salida}")
            os.startfile(ruta_salida)
            messagebox.showinfo("Proceso completado", f"Consolidación completada. Archivo guardado en:\n{ruta_salida}")
        except Exception as e:
            print(f"Error al guardar el archivo: {e}")
            messagebox.showerror("Error", f"Error al guardar el archivo: {e}")
    else:
        print("No se pudo consolidar ninguna hoja. Verifica que las hojas contengan la columna clave.")
        messagebox.showerror("Error", "No se pudo consolidar ninguna hoja.")

def cargar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entry_excel_path.delete(0, tk.END)
        entry_excel_path.insert(0, archivo)

def guardar_archivo():
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo:
        entry_salida_path.delete(0, tk.END)
        entry_salida_path.insert(0, archivo)

def ejecutar_consolidacion():
    excel_path = entry_excel_path.get()
    salida_path = entry_salida_path.get()

    if not excel_path or not salida_path:
        messagebox.showwarning("Advertencia", "Por favor, completa todos los campos.")
        return

    consolidar_hojas(excel_path, salida_path)

# Crear la interfaz de usuario
root = tk.Tk()
root.title("Consolidación de Hojas de Excel")

# Crear los widgets
label_excel_path = tk.Label(root, text="Ruta del archivo Excel:")
label_excel_path.grid(row=0, column=0, padx=10, pady=5, sticky="e")

entry_excel_path = tk.Entry(root, width=50)
entry_excel_path.grid(row=0, column=1, padx=10, pady=5)

button_cargar_excel = tk.Button(root, text="Cargar Excel", command=cargar_archivo)
button_cargar_excel.grid(row=0, column=2, padx=10, pady=5)

label_salida_path = tk.Label(root, text="Ruta de salida:")
label_salida_path.grid(row=2, column=0, padx=10, pady=5, sticky="e")

entry_salida_path = tk.Entry(root, width=50)
entry_salida_path.grid(row=2, column=1, padx=10, pady=5)

button_guardar_salida = tk.Button(root, text="Guardar archivo", command=guardar_archivo)
button_guardar_salida.grid(row=2, column=2, padx=10, pady=5)

button_ejecutar = tk.Button(root, text="Ejecutar consolidación", command=ejecutar_consolidacion)
button_ejecutar.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

# Iniciar la interfaz
root.mainloop()
