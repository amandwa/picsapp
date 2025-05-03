import os
import pandas as pd
import re
from tkinter import Tk, Canvas, Label, Button, filedialog, Entry, StringVar, OptionMenu
from zipfile import ZipFile
import threading
import shutil
from PIL import Image, ImageTk
import time

# ========== CONSTANTES ==========
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
COR_BARRA_TITULO = "#2D2D57"
FONTE_PADRAO = ("Segoe UI", 10)
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
caminho_pasta = "E:\\"

# ========== FUNÇÕES ==========
def limpar_termo(termo):
    return re.sub(r'[^a-zA-Z0-9]', '', termo)

def limpar_campos():
    global arquivos_nao_encontrados, df_planilha, option_menu
    entry_excel.delete(0, "end")
    coluna_selecionada.set("")
    status_var.set("")
    arquivos_nao_encontrados = []
    df_planilha = None
    if option_menu:
        option_menu.destroy()
        option_menu = None

def carregar_colunas():
    global df_planilha, option_menu
    caminho_excel = entry_excel.get()
    if not os.path.exists(caminho_excel):
        return
    try:
        df_planilha = pd.read_excel(caminho_excel, dtype=str)
        colunas = list(df_planilha.columns)
        coluna_selecionada.set(colunas[0] if colunas else "")
        if option_menu:
            option_menu.destroy()
        option_menu = OptionMenu(root, coluna_selecionada, *colunas)
        option_menu.place(x=250, y=250)
        Label(root, text="Selecione a coluna para nomear as fotos:", bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO).place(x=250, y=220)
    except:
        status_label.config(text="Erro ao carregar colunas.")

def gerar_zip():
    global rodando, NUM_ROWS, NUM_FOTOS_ENCONTRADAS, arquivos_nao_encontrados
    arquivos_nao_encontrados = []
    NUM_FOTOS_ENCONTRADAS = 0
    caminho_excel = entry_excel.get()
    if not os.path.exists(caminho_excel) or not os.path.exists(caminho_pasta):
        status_var.set("Caminho inválido.")
        rodando = False
        return
    try:
        df = pd.read_excel(caminho_excel, dtype=str)
    except:
        status_var.set("Erro ao ler planilha.")
        rodando = False
        return
    coluna = coluna_selecionada.get()
    if coluna not in df.columns:
        status_var.set("Coluna inválida.")
        rodando = False
        return
    
    # Aqui ficaria a implementação completa da geração do ZIP
    # Simulação de término do processo
    time.sleep(1)
    status_var.set("Processamento finalizado.")
    rodando = False

def iniciar_processamento():
    global rodando
    if rodando:
        return
    rodando = True
    threading.Thread(target=animar_bolinha, daemon=True).start()
    threading.Thread(target=gerar_zip, daemon=True).start()

# ========== FUNÇÕES DE INTERFACE ==========
def criar_entry_round(master, x, y, width=50):
    entry = Entry(master, width=width, font=FONTE_PADRAO, relief="flat", bg="white")
    entry.place(x=x, y=y, width=width)
    return entry

def criar_botao_round(master, texto, comando, x, y, largura=220, altura=40, cor=COR_BOTAO, cor_texto=COR_BOTAO_TEXTO):
    canvas = Canvas(master, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=cor, outline=cor)
    canvas.create_oval(largura-20, 0, largura, altura, fill=cor, outline=cor)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=cor, outline=cor)
    botao = Button(master, text=texto, command=comando, bg=cor, fg=cor_texto, font=FONTE_PADRAO, relief="flat", bd=0, activebackground=cor)
    botao.place(x=x + 10, y=y + 8)
    return botao

def animar_bolinha():
    global bolinha_pos, indo_para_direita
    if not rodando:
        bolinha_canvas.delete("all")
        return
    
    if indo_para_direita:
        bolinha_pos += 4
        if bolinha_pos > 80:
            indo_para_direita = False
    else:
        bolinha_pos -= 4
        if bolinha_pos < 0:
            indo_para_direita = True
    
    bolinha_canvas.delete("all")
    bolinha_canvas.create_oval(10 + bolinha_pos, 10, 22 + bolinha_pos, 22, fill=COR_BOTAO, outline=COR_BOTAO)
    bolinha_canvas.after(30, animar_bolinha)

# ========== CONFIGURAÇÃO DA INTERFACE ==========
root = Tk()
root.title("Buscador de Imagens por Excel")
root.geometry("650x400")
root.configure(bg=COR_FUNDO_APP)
root.resizable(False, False)
root.overrideredirect(True)

# Variáveis globais
rodando = False
NUM_ROWS = 0
NUM_FOTOS_ENCONTRADAS = 0
arquivos_nao_encontrados = []
df_planilha = None
option_menu = None
bolinha_pos = 0
indo_para_direita = True

# Barra superior
barra_superior = Canvas(root, width=650, height=35, bg=COR_BARRA_TITULO, highlightthickness=0)
barra_superior.place(x=0, y=0)

def iniciar_mover(event):
    root.x = event.x
    root.y = event.y

def mover_janela(event):
    x = event.x_root - root.x
    y = event.y_root - root.y
    root.geometry(f"+{x}+{y}")

barra_superior.bind("<Button-1>", iniciar_mover)
barra_superior.bind("<B1-Motion>", mover_janela)

botao_fechar = Button(barra_superior, text="X", command=lambda: root.quit(), bg=COR_BARRA_TITULO, fg=COR_TEXTO_FORTE,
                      font=("Arial", 12, "bold"), relief="flat", bd=0)
botao_fechar.place(x=615, y=5)

# Logo Altenburg
try:
    logo_path = os.path.join(os.path.dirname(__file__), "altenburg.png")
    img_original = Image.open(logo_path).resize((210, 210), Image.LANCZOS)
    logo_img = ImageTk.PhotoImage(img_original)
    logo_label = Label(root, image=logo_img, bg=COR_FUNDO_APP, borderwidth=0)
    logo_label.image = logo_img
    logo_label.place(x=10, y=40)
except Exception as e:
    print(f"Erro ao carregar logo: {e}")

# Variáveis de interface
coluna_selecionada = StringVar()
status_var = StringVar()

# Componentes da interface
Label(root, text="📄 Caminho do arquivo Excel:", bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE, 
      font=FONTE_PADRAO, anchor="w").place(x=250, y=160, width=300)

entry_excel = Entry(root, font=FONTE_PADRAO, relief="flat", bg="white")
entry_excel.place(x=250, y=190, width=300)

criar_botao_round(
    root, "     📂  Selecionar Arquivo",
    lambda: [
        entry_excel.delete(0, "end"),
        entry_excel.insert(0, filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])),
        carregar_colunas()
    ],
    x=350, y=230
)

Button(root, text="🗑️", command=limpar_campos, bg=COR_FUNDO_APP, fg="white", 
       font=FONTE_PADRAO, relief="flat", bd=0, activebackground=COR_FUNDO_APP).place(x=590, y=230, width=50, height=40)

gerar_zip_btn = criar_botao_round(root, "     🧾  Gerar ZIP com imagens", iniciar_processamento, x=350, y=280)

# Animação da bolinha
bolinha_canvas = Canvas(root, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
bolinha_canvas.place(x=350, y=320)

# Status
status_label = Label(root, textvariable=status_var, bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE, font=FONTE_PADRAO)
status_label.place(x=250, y=340)

# Créditos
creditos = Label(root, text=" © DEVELOPED BY AMANDA PRADO", bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE, font=("Segoe UI", 7))
creditos.place(x=320, y=385, anchor="n")

root.mainloop()
