import openpyxl

# Archivos de entrada y salida
input_file = "C:/Users/Ian/Documents/Python_scripts/Layouts/log-display_elabel_brief.txt"
output_file = "elabel_brief.xlsx"

# Extraer solo el bloque de "display elabel brief"
with open(input_file, "r", encoding="utf-8") as f:
    lines = f.readlines()

start_cmd = "display elabel brief"
capture = False
dash_count = 0
block = []

for line in lines:
    if start_cmd in line:
        capture = True
    if capture:
        block.append(line)
        # Detecta línea compuesta solo de guiones (punteada)
        if set(line.strip()) == {"-"}:
            dash_count += 1
            if dash_count == 3:  # tercera línea de guiones = fin del bloque
                capture = False
                break  # dejamos de leer después de la sección

# Crear libro de Excel
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Elabel Brief"

# Escribir encabezados
sheet.append(["Slot", "Description"])

with open(input_file, "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        # Saltar líneas vacías, separadores o encabezados
        if not line or set(line) == {"-"} or line.startswith("<") or "Slot" in line or "Elabel" in line:
            continue
        
        parts = line.split()
        if len(parts) >= 4:
            # Caso LPU / PIC / MPU / SFU / PWR / FAN
            if parts[0] in ["LPU", "PIC", "MPU", "SFU", "PWR", "FAN"]:
                slot = parts[0] + parts[1]   # Ej: LPU1, PIC0
                description = " ".join(parts[4:])  # Ignorar secciones BoardType y BarCode
            else:
                # Caso PEM0 / PEM1 → no tienen número separado
                slot = parts[0]
                description = " ".join(parts[2:])  # Ignorar secciones BoardType y BarCode
            
            # Guardar fila solo si hay descripción
            sheet.append([slot, description])

# Guardar Excel
workbook.save(output_file)
print(f"✅ Archivo Excel generado: {output_file}")

