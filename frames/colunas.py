from tkinter import Frame, Canvas, Label, Button, filedialog, StringVar, Entry
import pandas as pd
import os
import sys
from PIL import Image, ImageTk

# Cores e fontes
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
FONTE_PADRAO = ("Segoe UI", 10)

# Função para carregar arquivos corretamente em ambiente compilado
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Usado pelo PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ColunasFrame(Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=COR_FUNDO_APP, width=600, height=400)
        self.controller = controller
        self.pack_propagate(False)

        self.caminho_var = StringVar()
        self.status_var = StringVar()
        self.carregando_canvas = None
        self.after_id = None
        self.bolinha_pos = 0
        self.indo_para_direita = True

        self.criar_interface()

    def criar_interface(self):
        self.criar_logo()
        self.criar_campos()
        self.criar_botoes()
        self.criar_rodape()

        botao_casinha = Button(self, text="🏠", command=lambda: self.controller.show_frame("MenuFrame"),
                               bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO,
                               relief="flat", bd=0, activebackground=COR_FUNDO_APP,
                               activeforeground="white", highlightthickness=0)
        botao_casinha.place(x=-8, y=1, width=50, height=40)

        botao_lixeira = Button(self, text="🗑️", command=self.limpar_campos,
                               bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO,
                               relief="flat", bd=0, activebackground=COR_FUNDO_APP,
                               activeforeground="white", highlightthickness=0)
        botao_lixeira.place(x=460, y=245, width=50, height=40)

    def criar_logo(self):
        try:
            logo_path = resource_path("assets/altenburg.png")
            img_original = Image.open(logo_path).resize((210, 210), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img_original)
            logo_label = Label(self, image=self.logo_img, bg=COR_FUNDO_APP, borderwidth=0)
            logo_label.image = self.logo_img
            logo_label.place(x=10, y=18)
        except Exception:
            # Falha silenciosa ao carregar a imagem
            pass

        Label(self, text="| CELLY - Separador de Colunas",
              fg=COR_TEXTO_FORTE, bg=COR_FUNDO_APP, font=("Segoe UI Semibold", 13)).place(x=210, y=100)

    def criar_campos(self):
        Label(self, text="✂️ Separador (ex: , ou /):", bg=COR_FUNDO_APP,
              fg=COR_TEXTO_FORTE, font=FONTE_PADRAO).place(x=190, y=181)

        self.entrada_separador = Entry(self, width=5, font=FONTE_PADRAO,
                                       bg="#eee", fg="black", justify="center")
        self.entrada_separador.insert(0, ";")
        self.entrada_separador.place(x=380, y=185, height=22)

        Label(self, textvariable=self.status_var, bg=COR_FUNDO_APP,
              fg=COR_TEXTO_FORTE, font=("Verdana", 10, "italic")).place(relx=0.5, y=350, anchor="center")

    def criar_botoes(self):
        self.criar_botao_round("📂  Selecionar Planilha Excel", self.selecionar_planilha, 180, 215)
        self.criar_botao_round("📝  Gerar Arquivo Editado", self.mostrar_carregando, 180, 270)

    def criar_botao_round(self, texto, comando, x, y, largura=240, altura=40):
        canvas = Canvas(self, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
        canvas.place(x=x, y=y)
        canvas.create_oval(0, 0, 20, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        canvas.create_oval(largura-20, 0, largura, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        canvas.create_rectangle(10, 0, largura-10, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        botao = Button(self, text=texto, command=comando,
                       bg=COR_BOTAO, fg=COR_BOTAO_TEXTO, font=FONTE_PADRAO,
                       relief="flat", bd=0, activebackground=COR_BOTAO)
        botao.place(x=x + 10, y=y + 8)

    def criar_rodape(self):
        Label(self, text="© DEVELOPED BY AMANDA PRADO",
              bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE,
              font=("Segoe UI", 7)).place(relx=0.5, rely=1.0, anchor="s")

    def limpar_campos(self):
        self.entrada_separador.delete(0, "end")
        self.entrada_separador.insert(0, ";")
        self.caminho_var.set("")
        self.status_var.set("")

    def selecionar_planilha(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if caminho:
            self.caminho_var.set(caminho)
            self.status_var.set("Planilha selecionada.")

    def animar_bolinha(self):
        if not self.carregando_canvas:
            return

        if self.indo_para_direita:
            self.bolinha_pos += 4
            if self.bolinha_pos > 80:
                self.indo_para_direita = False
        else:
            self.bolinha_pos -= 4
            if self.bolinha_pos < 0:
                self.indo_para_direita = True

        self.carregando_canvas.coords(self.bolinha, 10 + self.bolinha_pos, 10, 22 + self.bolinha_pos, 22)
        self.after_id = self.after(30, self.animar_bolinha)

    def mostrar_carregando(self):
        if not self.caminho_var.get():
            self.status_var.set("Selecione um arquivo primeiro!")
            return

        self.status_var.set("")
        self.bolinha_pos = 0
        self.indo_para_direita = True

        self.carregando_canvas = Canvas(self, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
        self.carregando_canvas.place(x=375, y=320)
        self.bolinha = self.carregando_canvas.create_oval(10, 10, 22, 22, fill=COR_BOTAO, outline=COR_BOTAO)

        self.animar_bolinha()
        self.after(100, self.processar_excel)

    def processar_excel(self):
        try:
            caminho = self.caminho_var.get()
            sep = self.entrada_separador.get()

            df = pd.read_excel(caminho)

            col_solicitacoes = None
            col_quantidades = None

            for col in df.columns:
                if "solicita" in col.lower():
                    col_solicitacoes = col
                elif "quantidade" in col.lower():
                    col_quantidades = col

            if not col_solicitacoes or not col_quantidades:
                raise ValueError("Colunas 'Solicitações' ou 'Quantidade' não encontradas.")

            novas_linhas = []

            for _, row in df.iterrows():
                solicitacoes = str(row[col_solicitacoes]).split(sep)
                quantidades = str(row[col_quantidades]).split(sep)

                if len(solicitacoes) != len(quantidades):
                    raise ValueError("Quantidade de itens e quantidades não coincidem na linha.")

                for solicitacao, qtd in zip(solicitacoes, quantidades):
                    nova_linha = row.copy()
                    nova_linha[col_solicitacoes] = solicitacao.strip()
                    nova_linha[col_quantidades] = qtd.strip()
                    novas_linhas.append(nova_linha)

            novo_df = pd.DataFrame(novas_linhas)

            pasta = os.path.dirname(caminho)
            novo_caminho = os.path.join(pasta, "alteracao_ticket.xlsx")
            novo_df.to_excel(novo_caminho, index=False)

            self.status_var.set("Arquivo salvo como 'alteracao_ticket.xlsx'.")

        except Exception as e:
            self.status_var.set(f"Erro: {str(e)}")

        finally:
            if self.carregando_canvas:
                self.carregando_canvas.destroy()
                self.carregando_canvas = None

            if self.after_id:
                self.after_cancel(self.after_id)
                self.after_id = None
