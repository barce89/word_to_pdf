import os
from pathlib import Path
from docx2pdf import convert
import comtypes.client

def convertir_doc_a_pdf(input_path, output_path):
    """Convierte archivos .doc antiguos usando Word en segundo plano."""
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(input_path))
    doc.SaveAs(str(output_path), FileFormat=17) # 17 representa el formato PDF en Word
    doc.Close()
    word.Quit()

def main():
    # Obtiene la ruta actual donde se encuentra ejecutándose el script o .exe
    carpeta_actual = Path.cwd()
    nombre_carpeta = carpeta_actual.name
    
    # Define la carpeta de salida (nombre de la carpeta actual + "_PDFs")
    carpeta_salida = carpeta_actual / f"{nombre_carpeta}_PDFs"
    carpeta_salida.mkdir(exist_ok=True)
    
    print(f"Buscando archivos de Word en: {carpeta_actual}")
    print(f"Los PDFs se guardarán en: {carpeta_salida}\n")
    
    # Extensiones compatibles
    archivos_word = list(carpeta_actual.glob("*.docx")) + list(carpeta_actual.glob("*.doc"))
    
    # Filtrar para no procesar archivos temporales de Word que empiezan con ~$
    archivos_word = [f for f in archivos_word if not f.name.startswith("~$")]
    
    if not archivos_word:
        print("No se encontraron archivos .doc o .docx en esta carpeta.")
        input("\nPresiona Enter para salir...")
        return

    for archivo in archivos_word:
        extension = archivo.suffix.lower()
        nombre_salida = carpeta_salida / f"{archivo.stem}.pdf"
        
        try:
            print(f"Convirtiendo: {archivo.name}...")
            if extension == ".docx":
                # Conversión rápida para .docx
                convert(str(archivo), str(nombre_salida))
            elif extension == ".doc":
                # Conversión compatible para .doc antiguos
                convertir_doc_a_pdf(archivo, nombre_salida)
            print(f" -> ¡Éxito! Guardado como {nombre_salida.name}")
        except Exception as e:
            print(f" -> Error al convertir {archivo.name}: {e}")

    print("\n¡Proceso de conversión finalizado!")
    input("Presiona Enter para cerrar la ventana...")

if __name__ == "__main__":
    main()