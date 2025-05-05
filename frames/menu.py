from tkinter import Frame, Canvas, Label, Button, StringVar
from PIL import Image, ImageTk
import os
import sys

# Cores e fontes
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
FONTE_PADRAO = ("Segoe UI", 10)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class MenuFrame(Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=COR_FUNDO_APP, width=600, height=400)
        self.controller = controller
        self.pack_propagate(False)

        self.bolinha_pos = 0
        self.indo_para_direita = True
        self.carregando_canvas = None

        self.criar_logo()
        self.criar_interface()

    def criar_logo(self):
        try:
            logo_path = resource_path("assets/altenburg.png")
            img_original = Image.open(logo_path).resize((210, 210), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img_original)
            logo_label = Label(self, image=self.logo_img, bg=COR_FUNDO_APP, borderwidth=0)
            logo_label.image = self.logo_img
            logo_label.place(x=10, y=18)

            Label(self, text="| CELLY - Utilitários", bg=COR_FUNDO_APP,
                  fg=COR_TEXTO_FORTE, font=("Segoe UI Semibold", 13)).place(x=210, y=100)
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")

    def criar_interface(self):
        self.status_var = StringVar()
        Label(self, textvariable=self.status_var, bg=COR_FUNDO_APP,
              fg=COR_TEXTO_FORTE, font=("Verdana", 10, "italic")).place(x=215, y=325)

        self.criar_botao_round("     🔍  Buscador de Fotos", lambda: self.iniciar_transicao("ImagensFrame"), 190, 205)
        self.criar_botao_round("       📈  Separador Excel", lambda: self.iniciar_transicao("ColunasFrame"), 190, 255)

        Label(self, text=" © DEVELOPED BY AMANDA PRADO", bg=COR_FUNDO_APP,
              fg=COR_TEXTO_FORTE, font=("Segoe UI", 7)).place(x=300, y=385, anchor="center")

    def criar_botao_round(self, texto, comando, x, y, largura=220, altura=40):
        canvas = Canvas(self, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
        canvas.place(x=x, y=y)
        canvas.create_oval(0, 0, 20, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        canvas.create_oval(largura-20, 0, largura, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        canvas.create_rectangle(10, 0, largura-10, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        botao = Button(self, text=texto, command=comando, bg=COR_BOTAO, fg=COR_BOTAO_TEXTO,
                       font=FONTE_PADRAO, relief="flat", bd=0, activebackground=COR_BOTAO)
        botao.place(x=x + 10, y=y + 8)

    def animar_bolinha(self):
        if self.indo_para_direita:
            self.bolinha_pos += 4
            if self.bolinha_pos > 80:
                self.indo_para_direita = False
        else:
            self.bolinha_pos -= 4
            if self.bolinha_pos < 0:
                self.indo_para_direita = True

        if self.carregando_canvas and self.bolinha:
            self.carregando_canvas.coords(self.bolinha, 10 + self.bolinha_pos, 10, 22 + self.bolinha_pos, 22)
            self.after(30, self.animar_bolinha)

    def iniciar_transicao(self, nome_frame):
        self.status_var.set("")
        self.bolinha_pos = 0
        self.indo_para_direita = True

        self.carregando_canvas = Canvas(self, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
        self.carregando_canvas.place(x=245, y=300)
        self.bolinha = self.carregando_canvas.create_oval(10, 10, 22, 22, fill=COR_BOTAO, outline=COR_BOTAO)

        self.animar_bolinha()
        self.after(1000, lambda: self.finalizar_transicao(nome_frame))

    def finalizar_transicao(self, nome_frame):
        if self.carregando_canvas:
            self.carregando_canvas.destroy()
            self.carregando_canvas = None
        self.controller.show_frame(nome_frame)

    def tkraise(self, aboveThis=None):
        super().tkraise(aboveThis)
        if self.carregando_canvas:
            self.carregando_canvas.destroy()
            self.carregando_canvas = None
        self.status_var.set("")
