import os
import pandas as pd
import re
from tkinter import Tk, Canvas, Label, Button, filedialog, Entry
from zipfile import ZipFile
import threading
import shutil

# Constantes
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
caminho_pasta = "E:\\"  # Pasta fixa (ajuste conforme necessário)

# Variáveis globais
rodando = False
NUM_ROWS = 0
NUM_FOTOS_ENCONTRADAS = 0
arquivos_nao_encontrados = []

# Funções utilitárias
def limpar_termo(termo):
    return re.sub(r'[^a-zA-Z0-9]', '', termo)

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

def criar_botao_round(master, texto, comando, x, y, largura=180, altura=40, cor="#2e2e2e", cor_texto="white"):
    canvas = Canvas(master, width=largura, height=altura, bg="#f4e8d4", highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=cor, outline=cor)
    canvas.create_oval(largura-20, 0, largura, altura, fill=cor, outline=cor)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=cor, outline=cor)
    botao = Button(master, text=texto, command=comando, bg=cor, fg=cor_texto, font=("Verdana", 10), relief="flat", bd=0, activebackground=cor)
    botao.place(x=x + 5, y=y + 8)
    return botao

# Função de animação de bolinha suave
def animar_bolinha():
    raio = 10
    x = 0
    direcao = 1

    def mover():
        nonlocal x, direcao
        if not rodando:
            bolinha_canvas.delete("all")
            return
        bolinha_canvas.delete("all")
        bolinha_canvas.create_oval(10 + x, 10, 10 + 2 * raio + x, 10 + 2 * raio, fill="#888888", outline="")
        x += direcao * 2
        if x > 60 or x < 0:
            direcao *= -1
        bolinha_canvas.after(20, mover)

    mover()

# Função principal
def gerar_zip():
    global rodando, NUM_ROWS, NUM_FOTOS_ENCONTRADAS, arquivos_nao_encontrados
    arquivos_nao_encontrados = []
    NUM_FOTOS_ENCONTRADAS = 0

    caminho_excel = entry_excel.get()
    if not os.path.exists(caminho_excel) or not os.path.exists(caminho_pasta):
        status_label.config(text="Caminho inválido.")
        rodando = False
        return

    try:
        df = pd.read_excel(caminho_excel, dtype=str)
    except:
        status_label.config(text="Erro ao ler planilha.")
        rodando = False
        return

    termos = set()
    for _, row in df.iloc[1:].iterrows():  # Pula a primeira linha (cabeçalho)
        for val in row:
            if pd.notna(val):
                termo = limpar_termo(str(val).strip().lower())
                if termo:
                    termos.add(termo)

    NUM_ROWS = len(termos)
    encontrados = []
    encontrados_set = set()

    for root_dir, _, files in os.walk(caminho_pasta):
        for file in files:
            nome_base = os.path.splitext(file)[0].lower()
            nome_limpo = limpar_termo(nome_base)
            for termo in termos:
                if termo in nome_limpo:
                    encontrados.append(os.path.join(root_dir, file))
                    encontrados_set.add(termo)
                    break

    NUM_FOTOS_ENCONTRADAS = len(encontrados_set)
    arquivos_nao_encontrados = sorted(list(termos - encontrados_set))

    zip_path = os.path.join(DOWNLOAD_PATH, "catalogo.zip")
    count = 0

    with ZipFile(zip_path, 'w') as zipf:
        for file in encontrados:
            zipf.write(file, os.path.basename(file))
            count += 1
            status_label.config(text=f"{count} arquivos adicionados ao ZIP... ({NUM_FOTOS_ENCONTRADAS}/{NUM_ROWS})")
            root.update()

    rodando = False

    if arquivos_nao_encontrados:
        status_label.config(text=f"Concluído. {NUM_FOTOS_ENCONTRADAS} encontrados de {NUM_ROWS}.")
    else:
        status_label.config(text="Sucesso! Todos os arquivos foram localizados e zipados.")

def iniciar_processamento():
    global rodando
    rodando = True
    threading.Thread(target=animar_bolinha, daemon=True).start()
    threading.Thread(target=gerar_zip, daemon=True).start()

# Interface gráfica
root = Tk()
root.title("Buscador de Imagens por Excel")
root.geometry("610x460")
root.configure(bg="#f4e8d4")
root.resizable(False, False)

Label(root, text="Caminho do arquivo:", bg="#f4e8d4", font=("Verdana", 10)).place(x=14, y=160)
entry_excel = criar_entry_round(root, 20, 190)

criar_botao_round(root, "Selecionar Arquivo", lambda: [
    entry_excel.delete(0, "end"),
    entry_excel.insert(0, filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])),
], x=230, y=250, largura=140, altura=35, cor="#e0e0e0", cor_texto="black")

criar_botao_round(root, "Gerar ZIP com imagens", iniciar_processamento, x=210, y=310)

# Bolinha animada
bolinha_canvas = Canvas(root, width=100, height=40, bg="#f4e8d4", highlightthickness=0)
bolinha_canvas.place(relx=0.5, y=370, anchor="center")

# Status
status_label = Label(root, text="", bg="#f4e8d4", font=("Verdana", 10))
status_label.place(relx=0.5, y=420, anchor="center")

root.mainloop()
