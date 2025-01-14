import pandas as pd

def consolidar_planillas(rutas_planillas, clave_union, ruta_salida):
    """
    Consolida varias planillas en una sola con base en una clave de unión.

    :param rutas_planillas: Lista de rutas de los archivos Excel a consolidar.
    :param clave_union: Columna clave para unir las planillas.
    :param ruta_salida: Ruta del archivo Excel final consolidado.
    """
    planilla_final = None

    for ruta in rutas_planillas:
        # Carga la planilla
        planilla = pd.read_excel(ruta)
        
        # Si es la primera planilla, inicializa la planilla final
        if planilla_final is None:
            planilla_final = planilla
        else:
            # Realiza la unión con la planilla acumulada
            planilla_final = planilla_final.merge(planilla, on=clave_union, how='outer')

    # Exporta la planilla final consolidada
    planilla_final.to_excel(ruta_salida, index=False)
    print(f"Consolidación completa. Archivo final guardado en: {ruta_salida}")

# Ejemplo de uso:

# Rutas de las 5 planillas
rutas_planillas = [
    "planilla1.xlsx",  # Ruta del primer archivo Excel
    "planilla2.xlsx",  # Ruta del segundo archivo Excel
    "planilla3.xlsx",  # Ruta del tercer archivo Excel
    "planilla4.xlsx",  # Ruta del cuarto archivo Excel
    "planilla5.xlsx"   # Ruta del quinto archivo Excel
]

# Columna clave para unir las planillas
clave_union = "OC"  # Reemplazar con la clave real que vincula las planillas

# Ruta de la planilla final consolidada
ruta_salida = "planilla_final.xlsx"

# Ejecutar la consolidación
consolidar_planillas(rutas_planillas, clave_union, ruta_salida)