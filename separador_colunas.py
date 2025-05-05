from tkinter import Tk, Canvas, Label, Button, filedialog, StringVar, Entry
import pandas as pd
import os
from PIL import Image, ImageTk
import sys
import subprocess

# Cores e fontes
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
COR_BARRA_TITULO = "#2D2D57"
FONTE_PADRAO = ("Segoe UI", 10)

# Janela principal
root = Tk()
root.title("Separador de Colunas")
root.geometry("650x400")
root.configure(bg=COR_FUNDO_APP)
root.resizable(False, False)
root.overrideredirect(True)

# Variáveis globais
carregando_canvas = None
botao_gerar = None
after_id = None

# Ícone
try:
    icone_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    root.iconbitmap(icone_path)
except Exception as e:
    print(f"Erro ao carregar ícone: {e}")

# Barra superior
def iniciar_mover(event):
    root.x = event.x
    root.y = event.y

def mover_janela(event):
    x = event.x_root - root.x
    y = event.y_root - root.y
    root.geometry(f"+{x}+{y}")

barra_superior = Canvas(root, width=650, height=35, bg=COR_BARRA_TITULO, highlightthickness=0)
barra_superior.place(x=0, y=0)
barra_superior.bind("<Button-1>", iniciar_mover)
barra_superior.bind("<B1-Motion>", mover_janela)

# Botão fechar
def fechar():
    root.quit()

botao_fechar = Button(barra_superior, text="X", command=fechar, bg=COR_BARRA_TITULO, fg=COR_TEXTO_FORTE,
                      font=("Arial", 12, "bold"), relief="flat", bd=0)
botao_fechar.place(x=615, y=5)

# Criar botão arredondado
def criar_botao_round(master, texto, comando, x, y, largura=240, altura=40, cor=COR_BOTAO, cor_texto=COR_BOTAO_TEXTO):
    canvas = Canvas(master, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=cor, outline=cor)
    canvas.create_oval(largura-20, 0, largura, altura, fill=cor, outline=cor)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=cor, outline=cor)
    botao = Button(master, text=texto, command=comando, bg=cor, fg=cor_texto, font=FONTE_PADRAO,
                   relief="flat", bd=0, activebackground=cor)
    botao.place(x=x + 10, y=y + 8)
    return botao

# Logo
try:
    logo_path = os.path.join(os.path.dirname(__file__), "altenburg.png")
    img_original = Image.open(logo_path).resize((210, 210), Image.LANCZOS)
    logo_img = ImageTk.PhotoImage(img_original)
    logo_label = Label(root, image=logo_img, bg=COR_FUNDO_APP, borderwidth=0)
    logo_label.image = logo_img
    logo_label.place(x=10, y=40)
except Exception as e:
    print(f"Erro ao carregar logo: {e}")

# Texto lateral
texto_logo = Label(
    root,
    text="| CELLY - Separador de Colunas",
    bg=COR_FUNDO_APP,
    fg=COR_TEXTO_FORTE,
    font=("Segoe UI Semibold", 13)
)
texto_logo.place(x=210, y=126)

# Campo separador
label_sep = Label(root, text="Separador:", bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE, font=FONTE_PADRAO)
label_sep.place(x=215, y=185)

entrada_separador = Entry(root, width=5, font=FONTE_PADRAO, bg="#eee", fg="black", justify="center", bd=1, relief="flat")
entrada_separador.insert(0, ";")
entrada_separador.place(x=295, y=185, height=22)

# Variáveis
caminho_var = StringVar()
status_var = StringVar()

# Botão lixeira para limpar campos
def limpar_campos():
    entrada_separador.delete(0, "end")
    entrada_separador.insert(0, "|")
    caminho_var.set("")
    status_var.set("")

botao_limpar = Button(
    root,
    text="🗑️",
    command=limpar_campos,
    bg=COR_FUNDO_APP,
    fg="white",
    font=FONTE_PADRAO,
    activebackground=COR_FUNDO_APP,
    activeforeground="white",
    relief="flat",
    bd=0,
    cursor="hand2",
    highlightthickness=0,
    takefocus=0
)
botao_limpar.place(x=455, y=245, width=50, height=40)


# Status
status_label = Label(root, textvariable=status_var, bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE, font=("Verdana", 10, "italic"))
status_label.place(relx=0.5, y=350, anchor="center")


# Animação
def animar_bolinha():
    global bolinha_pos, indo_para_direita, carregando_canvas, after_id
    if not carregando_canvas:
        return

    if indo_para_direita:
        bolinha_pos += 4
        if bolinha_pos > 80:
            indo_para_direita = False
    else:
        bolinha_pos -= 4
        if bolinha_pos < 0:
            indo_para_direita = True

    try:
        carregando_canvas.coords(bolinha, 10 + bolinha_pos, 10, 22 + bolinha_pos, 22)
        after_id = root.after(30, animar_bolinha)
    except Exception as e:
        print(f"Erro na animação: {e}")

def mostrar_carregando():
    global carregando_canvas, bolinha, bolinha_pos, indo_para_direita

    status_var.set("")
    bolinha_pos = 0
    indo_para_direita = True

    if botao_gerar:
        botao_gerar["state"] = "disabled"

    carregando_canvas = Canvas(root, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
    carregando_canvas.place(x=275, y=320)

    bolinha = carregando_canvas.create_oval(10, 10, 22, 22, fill=COR_BOTAO, outline=COR_BOTAO)
    animar_bolinha()

    root.after(100, processar_excel)

# Processar planilha
def processar_excel():
    global carregando_canvas, after_id
    try:
        caminho = caminho_var.get()
        if not caminho:
            status_var.set("Selecione um arquivo primeiro.")
            return

        sep = entrada_separador.get()
        if not sep:
            status_var.set("Informe o separador.")
            return

        df = pd.read_excel(caminho)

        if 'F' not in df.columns and df.columns.size >= 6:
            df.columns.values[5] = 'F'
        if 'G' not in df.columns and df.columns.size >= 7:
            df.columns.values[6] = 'G'

        for coluna in ['F', 'G']:
            if coluna in df.columns:
                df[[f'{coluna}_parte1', f'{coluna}_parte2']] = df[coluna].astype(str).str.split(pat=sep, n=1, expand=True)

        novo_nome = os.path.splitext(os.path.basename(caminho))[0] + "_SEPARADO.xlsx"
        novo_caminho = os.path.join(os.path.dirname(caminho), novo_nome)
        df.to_excel(novo_caminho, index=False)
        status_var.set(f"Planilha salva: {novo_nome}")

    except Exception as e:
        status_var.set(f"Erro: {e}")
    
    finally:
        if carregando_canvas:
            carregando_canvas.destroy()
        carregando_canvas = None
        if after_id:
            root.after_cancel(after_id)
            after_id = None
        if botao_gerar:
            botao_gerar["state"] = "normal"

# Botão para selecionar planilha
def selecionar_planilha():
    caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if caminho:
        caminho_var.set(caminho)
        status_var.set("Planilha selecionada.")

# Botão "Selecionar Planilha"
criar_botao_round(root, "     📂  Selecionar Planilha Excel", selecionar_planilha, x=215, y=215)

# Botão "Gerar Arquivo Editado"
botao_gerar = criar_botao_round(root, "     📝  Gerar Arquivo Editado", mostrar_carregando, x=215, y=270)

# Créditos
creditos = Label(
    root,
    text=" © DEVELOPED BY AMANDA PRADO",
    bg=COR_FUNDO_APP,
    fg=COR_TEXTO_FORTE,
    font=("Segoe UI", 7)
)
creditos.place(x=319, y=380, anchor="n")

def voltar_menu(nome_script="iniciar.py"):
    """
    Fecha o app atual e abre o menu principal (ou outro script especificado).
    """
    caminho_script = os.path.join(os.path.dirname(__file__), nome_script)
    subprocess.Popen([sys.executable, caminho_script])
    root.destroy()

botao_voltar = Button(
    root,
    text="🏠",
    command=voltar_menu,
    bg=COR_FUNDO_APP,
    fg="white",
    font=FONTE_PADRAO,
    activebackground=COR_FUNDO_APP,
    activeforeground="white",
    relief="flat",
    bd=0,
    cursor="hand2",
    highlightthickness=0,
    takefocus=0
)
botao_voltar.place(x=-5, y=36, width=50, height=40)

root.mainloop()
