import tkinter as tk
from tkinter import filedialog
from docx import Document
from bs4 import BeautifulSoup

def convert_to_html():
    # Abre un cuadro de diálogo para seleccionar el archivo de Word
    file_path = filedialog.askopenfilename(filetypes=[("Archivos de Word", "*.docx")])
    
    if not file_path:
        return  # Salir si el usuario cancela la selección de archivo

    # Abre el archivo de Word
    doc = Document(file_path)

    # Inicializa una cadena HTML
    html_output = "<html><head><title>Documento HTML</title></head><body>"

    # Recorre los párrafos en el documento de Word
    for paragraph in doc.paragraphs:
        # Agrega etiquetas de párrafo y texto al HTML
        html_output += f"<p>{paragraph.text}</p>"

    # Cierra la etiqueta body y html
    html_output += "</body></html>"

    # Abre un cuadro de diálogo para elegir la ubicación de guardado del archivo HTML
    save_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("Archivos HTML", "*.html")])

    if save_path:
        # Guarda el HTML en el archivo seleccionado
        with open(save_path, "w", encoding="utf-8") as html_file:
            html_file.write(html_output)

        print("Documento HTML generado con éxito en:", save_path)

# Crear la ventana principal
root = tk.Tk()
root.withdraw()  # Oculta la ventana principal

# Ejecuta la función para convertir a HTML
convert_to_html()

# Cierra la aplicación después de la conversión
root.destroy()
