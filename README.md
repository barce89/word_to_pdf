# word_to_pdf
Script hecho para pasar archivos de word a pdf

Convertidor Masivo de Word a PDF (.doc y .docx)

Herramienta ligera y portable desarrollada en Python para la conversión masiva de documentos de Microsoft Word (.doc y .docx) a formato PDF de forma automatizada dentro del directorio de ejecución.

🛠️ Requisitos del Sistema

Sistema Operativo: Windows 10 / 11 (64-bits).

Dependencia Clave: Microsoft Word debe estar instalado en la computadora donde se ejecute el archivo .exe (utiliza la automatización COM de Windows para preservar el formato original).

Entorno de Desarrollo (opcional para compilar): Python 3.12+ en Windows.

⚙️ Estructura del Proyecto

├── convertidor.py       # Script principal de Python
├── requirements.txt     # Dependencias del proyecto
├── python_version.txt   # Versión de Python utilizada
└── dist/
    └── convertidor.exe  # Ejecutable final listo para usar


🚀 Instalación y Configuración del Entorno (Para Desarrolladores)

Si deseas modificar el código o volver a compilar el ejecutable desde cero:

1. Crear el entorno virtual en Windows (PowerShell)

# Crear entorno virtual con Python 3.12/3.13
py -3.12 -m venv .venv

# Activar el entorno virtual
.\.venv\Scripts\Activate


2. Instalar dependencias

En caso de tener restricciones de certificados red/SSL, puedes utilizar los hosts de confianza:

pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
