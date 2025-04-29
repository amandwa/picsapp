import os
import pandas as pd  # Incluído, mas não usado diretamente. Pode ser usado para ler dados depois.
import re
from tkinter import Tk, Canvas, Label, Button, filedialog, Entry
from zipfile import ZipFile
import threading
import shutil

# Caminho para a pasta Downloads do usuário
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CATALOGO_PATTERN = re.compile(r"catalogo_\d+\.zip", re.IGNORECASE)

def processar_catalogos():
    arquivos_catalogo = [f for f in os.listdir(DOWNLOAD_PATH) if CATALOGO_PATTERN.match(f)]

    if not arquivos_catalogo:
        print("Nenhum arquivo catalogo_X.zip encontrado.")
        return

    temp_dir = os.path.join(DOWNLOAD_PATH, "TEMP_CATALOGOS")
    os.makedirs(temp_dir, exist_ok=True)

    # Extrai os arquivos
    for arquivo in arquivos_catalogo:
        caminho_zip = os.path.join(DOWNLOAD_PATH, arquivo)
        with ZipFile(caminho_zip, 'r') as zip_ref:
            extract_path = os.path.join(temp_dir, os.path.splitext(arquivo)[0])
            zip_ref.extractall(extract_path)
            print(f"Extraído: {arquivo}")

    # Cria um novo zip com todos os arquivos extraídos
    caminho_saida = os.path.join(DOWNLOAD_PATH, "CATALOGOS.zip")
    with ZipFile(caminho_saida, 'w') as zip_out:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zip_out.write(file_path, arcname)
                print(f"Adicionado ao CATALOGOS.zip: {arcname}")

    # Remove as pastas extraídas
    shutil.rmtree(temp_dir)
    print("Arquivos temporários removidos.")
    print(f"Arquivo final criado em: {caminho_saida}")

# Interface gráfica básica com Tkinter
def iniciar_interface():
    def on_processar_click():
        threading.Thread(target=processar_catalogos).start()

    root = Tk()
    root.title("Processar Catálogos")
    root.geometry("300x150")

    canvas = Canvas(root, height=150, width=300)
    canvas.pack()

    label = Label(root, text="Clique para processar catálogos na pasta Downloads")
    label.pack(pady=10)

    botao = Button(root, text="Processar", command=on_processar_click)
    botao.pack()

    root.mainloop()

# Executar
if __name__ == "__main__":
    iniciar_interface()

