import os
import pandas as pd
import re
from tkinter import Tk, Canvas, Label, Button, filedialog, Entry
from zipfile import ZipFile
import threading
import shutil

# Constantes
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CATALOGO_PATTERN = re.compile(r"catalogo_\d+\.zip", re.IGNORECASE)
MAX_ARQUIVOS_POR_ZIP = 300

# Função para limpar os termos
def limpar_termo(termo):
    return re.sub(r'[^a-zA-Z0-9]', '', termo)

# Função para criar Entry com fundo arredondado
def criar_entry_round(master, x, y, width=50):
    canvas = Canvas(master, width=550, height=30, bg="#f4e8d4", highlightthickness=0)
    canvas.place(x=x-5, y=y-5)
    canvas.create_oval(0, 0, 10, 10, fill="white", outline="white")
    canvas.create_oval(540, 0, 550, 10, fill="white", outline="white")
    canvas.create_oval(0, 20, 10, 30, fill="white", outline="white")
    canvas.create_oval(540, 20, 550, 30, fill="white", outline="white")
    canvas.create_rectangle(10, 0, 540, 30, fill="white", outline="white")
    entry = Entry(master, width=width, font=("Verdana", 10), relief="flat", bg="white")
    entry.place(x=x, y=y)
    return entry

# Criar botão com cantos arredondados
def criar_botao_round(master, texto, comando, x, y, largura=180, altura=40, cor="#2e2e2e", cor_texto="white"):
    canvas = Canvas(master, width=largura, height=altura, bg="#f4e8d4", highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=cor, outline=cor)
    canvas.create_oval(largura-20, 0, largura, altura, fill=cor, outline=cor)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=cor, outline=cor)
    botao = Button(master, text=texto, command=comando, bg=cor, fg=cor_texto, font=("Verdana", 10), relief="flat", bd=0, activebackground=cor)
    botao.place(x=x + 5, y=y + 8)
    return botao

# Spinner de progresso
progresso = 0
progresso_total = 100

def atualizar_barra():
    barra_canvas.delete("all")
    barra_canvas.create_rectangle(0, 0, 300, 20, fill="#ccc", outline="")
    barra_canvas.create_rectangle(0, 0, 3 * progresso, 20, fill="#4caf50", outline="")

# Gera os ZIPs com limite de arquivos
def gerar_zip():
    global progresso, progresso_total
    caminho_excel = entry_excel.get()
    caminho_pasta = "E:\\"

    if not os.path.exists(caminho_excel) or not os.path.exists(caminho_pasta):
        return

    try:
        df = pd.read_excel(caminho_excel, dtype=str)
    except:
        return

    termos = set()
    for _, row in df.iterrows():
        for val in row:
            if pd.notna(val):
                termo = limpar_termo(str(val).strip().lower())
                if termo:
                    termos.add(termo)

    encontrados = []
    for root_dir, _, files in os.walk(caminho_pasta):
        for file in files:
            nome_base = os.path.splitext(file)[0].lower()
            nome_limpo = limpar_termo(nome_base)
            for termo in termos:
                if termo in nome_limpo:
                    encontrados.append(os.path.join(root_dir, file))
                    break

    progresso_total = len(encontrados)
    progresso = 0

    for i in range(0, len(encontrados), MAX_ARQUIVOS_POR_ZIP):
        zip_path = os.path.join(DOWNLOAD_PATH, f"catalogo_{i//MAX_ARQUIVOS_POR_ZIP + 1}.zip")
        with ZipFile(zip_path, 'w') as zipf:
            for file in encontrados[i:i+MAX_ARQUIVOS_POR_ZIP]:
                zipf.write(file, os.path.basename(file))
                progresso += 1
                atualizar_barra()

# Junta todos os ZIPs criados em um único arquivo final
def processar_catalogos():
    arquivos_catalogo = [f for f in os.listdir(DOWNLOAD_PATH) if CATALOGO_PATTERN.match(f)]
    if not arquivos_catalogo:
        return

    temp_dir = os.path.join(DOWNLOAD_PATH, "TEMP_CATALOGOS")
    os.makedirs(temp_dir, exist_ok=True)

    for arquivo in arquivos_catalogo:
        caminho_zip = os.path.join(DOWNLOAD_PATH, arquivo)
        with ZipFile(caminho_zip, 'r') as zip_ref:
            extract_path = os.path.join(temp_dir, os.path.splitext(arquivo)[0])
            zip_ref.extractall(extract_path)

    caminho_saida = os.path.join(DOWNLOAD_PATH, "CATALOGOS.zip")
    with ZipFile(caminho_saida, 'w') as zip_out:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zip_out.write(file_path, arcname)

    shutil.rmtree(temp_dir)

    # Mensagem final
    barra_canvas.create_text(150, 10, text="Concluído! Arquivo 'Catalogos' salvo em sua pasta Downloads.", fill="black", font=("Verdana", 9))

# Executa tudo em thread

def executar_processo():
    gerar_zip()
    processar_catalogos()

# Interface
root = Tk()
root.title("Buscador de Imagens por Excel")
root.geometry("610x480")
root.configure(bg="#f4e8d4")
root.resizable(False, False)

Label(root, text="Caminho do arquivo:", bg="#f4e8d4", font=("Verdana", 10)).place(x=14, y=160)
entry_excel = criar_entry_round(root, 20, 190)

criar_botao_round(root, "Selecionar Arquivo", lambda: [
    entry_excel.delete(0, "end"),
    entry_excel.insert(0, filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")]))
], x=230, y=250, largura=140, altura=35, cor="#e0e0e0", cor_texto="black")

criar_botao_round(root, "Gerar ZIP com imagens", lambda: threading.Thread(target=executar_processo).start(), x=210, y=310)

barra_canvas = Canvas(root, width=300, height=20, bg="#f4e8d4", highlightthickness=0)
barra_canvas.place(relx=0.5, y=380, anchor="center")

root.mainloop()
