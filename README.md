# Convertidor Masivo de Word a PDF (.doc y .docx)

Herramienta ligera y portable desarrollada en Python para la conversión masiva de documentos de Microsoft Word (`.doc` y `.docx`) a formato PDF de forma automatizada dentro del directorio de ejecución.

---

## 🛠️ Requisitos del Sistema

- **Sistema Operativo:** Windows 10 / 11 (64-bits).
- **Dependencia Clave:** **Microsoft Word** debe estar instalado en la computadora donde se ejecute el archivo `.exe` (utiliza la automatización COM de Windows para preservar el formato original).
- **Entorno de Desarrollo (opcional para compilar):** Python 3.12+ en Windows.

---

## ⚙️ Estructura del Proyecto

```text
├── convertidor.py       # Script principal de Python
├── requirements.txt     # Dependencias del proyecto
├── python_version.txt   # Versión de Python utilizada
└── dist/
    └── convertidor.exe  # Ejecutable final listo para usar
```

---

## 🚀 Instalación y Configuración del Entorno (Para Desarrolladores)

Si deseas modificar el código o volver a compilar el ejecutable desde cero:

### 1. Crear el entorno virtual en Windows (PowerShell)
```powershell
# Crear entorno virtual con Python 3.12/3.13
py -3.12 -m venv .venv

# Activar el entorno virtual
.\.venv\Scripts\Activate
```

### 2. Instalar dependencias
En caso de tener restricciones de certificados red/SSL, puedes utilizar los hosts de confianza:

```powershell
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
```

---

## 📦 Compilación a Archivo Ejecutable (`.exe`)

Para empaquetar el script en un único ejecutable autónomo y portable sin requerir Python instalado:

```powershell
pyinstaller --onefile --hidden-import=comtypes convertidor.py
```

El ejecutable generado se guardará automáticamente en la carpeta **`dist/convertidor.exe`**.

---

## 💻 Uso del Ejecutable (`convertidor.exe`)

1. Copia o mueve el archivo **`convertidor.exe`** a la carpeta donde se encuentren tus documentos de Word (`.doc` o `.docx`).
2. Haz doble clic en **`convertidor.exe`**.
3. El programa detectará automáticamente todos los archivos compatibles en el directorio y creará una carpeta llamada **`_PDFs`** donde guardará las versiones convertidas.
4. Presiona **Enter** al finalizar para cerrar la ventana de la consola.

---

## 📄 Notas Adicionales

- El `.exe` es completamente independiente de Python y de las librerías de desarrollo.
- Asegúrate de cerrar los archivos de Word antes de ejecutar el conversor para evitar bloqueos de lectura/escritura.
