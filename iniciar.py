from tkinter import Tk, Canvas, Label, Button, StringVar
import subprocess
import sys
import os
from PIL import Image, ImageTk

# Cores e fontes
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
COR_BARRA_TITULO = "#2D2D57"
FONTE_PADRAO = ("Segoe UI", 10)

# Janela principal
root = Tk()
root.title("CELLY")
root.geometry("650x400")
root.configure(bg=COR_FUNDO_APP)
root.resizable(False, False)
root.overrideredirect(True)

# Ícone do app
try:
    icone_path = os.path.join(os.path.dirname(__file__), "icon.ico")
    root.iconbitmap(icone_path)
except Exception as e:
    print(f"Erro ao carregar ícone: {e}")

# Barra superior personalizada
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

# Botão de fechar customizado
def fechar():
    root.quit()

botao_fechar = Button(barra_superior, text="X", command=fechar, bg=COR_BARRA_TITULO, fg=COR_TEXTO_FORTE,
                      font=("Arial", 12, "bold"), relief="flat", bd=0)
botao_fechar.place(x=615, y=5)

# Criar botão arredondado
def criar_botao_round(master, texto, comando, x, y, largura=220, altura=40, cor=COR_BOTAO, cor_texto=COR_BOTAO_TEXTO):
    canvas = Canvas(master, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=cor, outline=cor)
    canvas.create_oval(largura-20, 0, largura, altura, fill=cor, outline=cor)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=cor, outline=cor)
    botao = Button(master, text=texto, command=comando, bg=cor, fg=cor_texto, font=FONTE_PADRAO,
                   relief="flat", bd=0, activebackground=cor)
    botao.place(x=x + 10, y=y + 8)
    return botao

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

# Texto ao lado do logo
texto_logo = Label(
    root,
    text="| CELLY - Facilitador",
    bg=COR_FUNDO_APP,
    fg=COR_TEXTO_FORTE,
    font=("Segoe UI Semibold", 13)
)
texto_logo.place(x=210, y=126)

# Status
status_var = StringVar()
status_label = Label(root, textvariable=status_var, bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE, font=("Verdana", 10, "italic"))
status_label.place(x=270, y=370)

# Animação da bolinha
def animar_bolinha():
    global bolinha_pos, indo_para_direita
    if indo_para_direita:
        bolinha_pos += 4
        if bolinha_pos > 80:
            indo_para_direita = False
    else:
        bolinha_pos -= 4
        if bolinha_pos < 0:
            indo_para_direita = True
    carregando_canvas.coords(bolinha, 10 + bolinha_pos, 10, 22 + bolinha_pos, 22)
    root.after(30, animar_bolinha)

# Ações dos botões
def mostrar_carregando(e_qual):
    global carregando_canvas, bolinha, bolinha_pos, indo_para_direita

    status_var.set("")
    bolinha_pos = 0
    indo_para_direita = True

    carregando_canvas = Canvas(root, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
    carregando_canvas.place(x=275, y=318)

    bolinha = carregando_canvas.create_oval(10, 10, 22, 22, fill=COR_BOTAO, outline=COR_BOTAO)

    animar_bolinha()
    root.update_idletasks()
    root.after(1500, lambda: abrir_app(e_qual))

def abrir_app(e_qual):
    script = "picsapp.py" if e_qual == "picsapp" else "separador_colunas.py"
    caminho_script = os.path.join(os.path.dirname(__file__), script)
    subprocess.Popen([sys.executable, caminho_script])
    root.destroy()

# Botões principais
criar_botao_round(root, "     🔍  Buscador de Fotos", lambda: mostrar_carregando("picsapp"), x=215, y=205)
criar_botao_round(root, "       📈  Separador Excel", lambda: mostrar_carregando("separador"), x=215, y=260)

# Créditos
creditos = Label(
    root,
    text=" © DEVELOPED BY AMANDA PRADO",
    bg=COR_FUNDO_APP,
    fg=COR_TEXTO_FORTE,
    font=("Segoe UI", 7)
)
creditos.place(x=319, y=380, anchor="n")

root.mainloop()
