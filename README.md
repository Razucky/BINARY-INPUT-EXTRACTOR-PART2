# Extractor de Entradas Binarias

Herramienta para extraer entradas binarias de planos de protección en PDF y exportarlas a Excel.

## 🚀 Descarga Rápida (Sin Python)

Ve a **Releases** (barra lateral derecha) y descarga:
- **Windows:** `BinaryInputExtractor.exe`
- **Mac:** `BinaryInputExtractor`
- **Linux:** `BinaryInputExtractor`

## 📋 Cómo Usar

1. Ejecuta el programa
2. Clic en **Examinar** para seleccionar hasta 3 archivos PDF
3. Elige dónde guardar el archivo Excel de salida
4. Clic en **Extraer Entradas Binarias**
5. Cada PDF se convierte en una pestaña separada en Excel

## 🔌 Dispositivos Soportados

- PCS-931S (NR Electric)
- SEL-411L (Schweitzer)
- PCS-9705S (NR Electric Bay Controller)
- UDF-506 (NR Electric)
- PCS-915SD (NR Electric Bus Protection)
- TESLA 4000 (ERL Power System Recorder)

## Columnas del Excel de Salida

| Columna | Descripción |
|---------|-------------|
| Substation | Nombre de la subestación |
| Bay | Bahía o línea |
| Voltage | Nivel de tensión |
| Switchgear | Tablero |
| Device | Tag del dispositivo (ej: -F01) |
| Model | Modelo del dispositivo |
| Function | Función del dispositivo |
| Board/Slot | Tarjeta o slot |
| Input_ID | ID de entrada (ej: BI_01) |
| Input_Number | Número de entrada |
| Description_Line1 | Primera línea de descripción |
| Description_Line2 | Segunda línea de descripción |
| Full_Description | Descripción completa |
| Page | Número de página del PDF |
