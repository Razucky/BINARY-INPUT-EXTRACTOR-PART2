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

---

# Instrucciones para Crear el .exe

## Paso 1: Crear cuenta en GitHub

1. Ve a https://github.com
2. Clic en "Sign up" (es gratis)

## Paso 2: Crear nuevo repositorio

1. Clic en el botón **+** (arriba a la derecha) → **New repository**
2. Nombre: `extractor-entradas-binarias`
3. Selecciona **Public** (requerido para builds gratis)
4. Marca **Add a README file**
5. Clic en **Create repository**

## Paso 3: Subir archivos

### Opción A: Subir binary_input_gui.py

1. En tu repositorio, clic en **Add file** → **Upload files**
2. Arrastra el archivo `binary_input_gui.py`
3. Clic en **Commit changes**

### Opción B: Crear el workflow (IMPORTANTE)

**GitHub no permite subir carpetas que empiecen con punto.** Debes crear el archivo manualmente:

1. Clic en **Add file** → **Create new file**
2. En el campo de nombre, escribe exactamente:
   ```
   .github/workflows/build.yml
   ```
3. Copia y pega el contenido del archivo `build.yml` (ver abajo)
4. Clic en **Commit changes**

## Paso 4: Esperar el build

1. Ve a la pestaña **Actions**
2. Verás "Build Executables" ejecutándose (punto amarillo)
3. Espera 3-5 minutos hasta que aparezca ✓ verde

## Paso 5: Descargar el .exe

1. Clic en **Releases** (barra lateral derecha)
2. Descarga `BinaryInputExtractor.exe`

---

## Contenido de build.yml

```yaml
name: Build Executables

on:
  push:
    branches: [ main ]
  workflow_dispatch:

jobs:
  build:
    strategy:
      matrix:
        include:
          - os: windows-latest
            name: Windows
            artifact: BinaryInputExtractor.exe
          - os: macos-latest
            name: macOS
            artifact: BinaryInputExtractor
          - os: ubuntu-latest
            name: Linux
            artifact: BinaryInputExtractor

    runs-on: ${{ matrix.os }}
    
    steps:
    - uses: actions/checkout@v4
    
    - name: Set up Python
      uses: actions/setup-python@v5
      with:
        python-version: '3.11'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pdfplumber openpyxl pyinstaller
    
    - name: Build executable
      run: |
        pyinstaller --onefile --windowed --name BinaryInputExtractor binary_input_gui.py
    
    - name: Upload artifact
      uses: actions/upload-artifact@v4
      with:
        name: BinaryInputExtractor-${{ matrix.name }}
        path: dist/${{ matrix.artifact }}
        retention-days: 90

  release:
    needs: build
    runs-on: ubuntu-latest
    if: github.event_name == 'push'
    
    steps:
    - name: Download all artifacts
      uses: actions/download-artifact@v4
      with:
        path: artifacts
    
    - name: Create Release
      uses: softprops/action-gh-release@v1
      with:
        tag_name: v1.0.${{ github.run_number }}
        name: Release v1.0.${{ github.run_number }}
        files: |
          artifacts/BinaryInputExtractor-Windows/BinaryInputExtractor.exe
          artifacts/BinaryInputExtractor-macOS/BinaryInputExtractor
          artifacts/BinaryInputExtractor-Linux/BinaryInputExtractor
        draft: false
        prerelease: false
      env:
        GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

---

## Solución de Problemas

### No veo la pestaña "Actions"
- Asegúrate de que el repositorio sea **Public**
- Ve a Settings → Actions → General → Selecciona "Allow all actions"

### El build falló
- Clic en el build fallido para ver el error
- Error común: error de sintaxis en el archivo workflow

### No veo Releases
- El primer release se crea automáticamente después de un build exitoso
- Espera a que el build termine completamente

---

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
