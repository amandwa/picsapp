import os
import pandas as pd
import re
from tkinter import *
from tkinter import filedialog
from zipfile import ZipFile
import threading
import shutil
from PIL import Image, ImageTk
import sys

# Cores e constantes
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
COR_BARRA_TITULO = "#2D2D57"
FONTE_PADRAO = ("Segoe UI", 10)
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CAMINHO_PASTA_IMAGENS = "E:\\"


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ImagensFrame(Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg=COR_FUNDO_APP)
        self.controller = controller
        self.rodando = False
        self.arquivos_nao_encontrados = []
        self.df_planilha = None
        self.option_menu = None
        self.bolinha_pos = 0
        self.indo_para_direita = True
        self.coluna_selecionada = StringVar(self)
        self.status_var = StringVar(self)
        self.logo_img = None  # Adicionado para armazenar a imagem

        self.criar_interface()

    def limpar_termo(self, termo):
        termo = str(termo).split('-')[0]
        return re.sub(r'[^0-9]', '', termo)

    def carregar_colunas(self):
        caminho_excel = self.entry_excel.get()
        if not os.path.exists(caminho_excel):
            self.status_var.set("Arquivo não encontrado!")
            return

        try:
            self.df_planilha = pd.read_excel(caminho_excel, dtype=str)
            colunas = list(self.df_planilha.columns)
            self.coluna_selecionada.set(colunas[0] if colunas else "")

            if self.option_menu:
                self.option_menu.destroy()

            self.option_menu = OptionMenu(self, self.coluna_selecionada, *colunas)
            self.option_menu.config(bg="white", font=FONTE_PADRAO, highlightthickness=0)
            self.option_menu.place(x=20, y=250)

            Label(self, text="Selecione a coluna 'Referência + Derivação':",
                  bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO).place(x=20, y=220)

            self.status_var.set("Planilha carregada com sucesso!")
        except Exception as e:
            self.status_var.set(f"Erro ao carregar planilha: {str(e)}")

    def gerar_zip(self):
        caminho_excel = self.entry_excel.get()
        if not os.path.exists(caminho_excel):
            self.status_var.set("Arquivo Excel não encontrado!")
            self.rodando = False
            return

        if not os.path.exists(CAMINHO_PASTA_IMAGENS):
            self.status_var.set(f"Pasta de imagens não encontrada em: {CAMINHO_PASTA_IMAGENS}")
            self.rodando = False
            return

        try:
            df = pd.read_excel(caminho_excel, dtype=str)
        except Exception as e:
            self.status_var.set(f"Erro ao ler planilha: {str(e)}")
            self.rodando = False
            return

        coluna = self.coluna_selecionada.get()
        if coluna not in df.columns:
            self.status_var.set("Coluna selecionada não existe na planilha!")
            self.rodando = False
            return

        termos_busca = {}
        for _, row in df.iterrows():
            valor = str(row[coluna]) if coluna in row and pd.notna(row[coluna]) else ""
            if valor:
                termo = self.limpar_termo(valor)
                if termo:
                    termos_busca[termo] = row[coluna]

        encontrados = []
        contador_nomes = {}

        for root_dir, _, files in os.walk(CAMINHO_PASTA_IMAGENS):
            for file in files:
                if not self.rodando:
                    break

                nome_arquivo = os.path.splitext(file)[0]
                nome_limpo = self.limpar_termo(nome_arquivo)

                for termo_busca, nome_final_base in termos_busca.items():
                    if termo_busca in nome_limpo:
                        ext = os.path.splitext(file)[1].lower()
                        nome_base_sem_derivacao = nome_final_base.split('-')[0]
                        contador = contador_nomes.get(nome_base_sem_derivacao, 0)
                        nome_final = f"{nome_base_sem_derivacao}_{contador}{ext}" if contador > 0 else f"{nome_base_sem_derivacao}{ext}"
                        contador_nomes[nome_base_sem_derivacao] = contador + 1

                        origem = os.path.join(root_dir, file)
                        destino = os.path.join(DOWNLOAD_PATH, nome_final)

                        try:
                            shutil.copy2(origem, destino)
                            encontrados.append(destino)
                            self.status_var.set(f"Encontrado: {nome_final}")
                            self.update()
                        except Exception as e:
                            print(f"Erro ao copiar {origem}: {str(e)}")

        if encontrados:
            zip_path = os.path.join(DOWNLOAD_PATH, "catalogo_imagens.zip")
            try:
                with ZipFile(zip_path, 'w') as zipf:
                    for arquivo in encontrados:
                        if os.path.exists(arquivo):
                            zipf.write(arquivo, os.path.basename(arquivo))
                            os.remove(arquivo)
                            self.status_var.set(f"Adicionando ao ZIP: {os.path.basename(arquivo)}")
                            self.update()
            except Exception as e:
                self.status_var.set(f"Erro ao criar ZIP: {str(e)}")
                self.rodando = False
                return

        termos_encontrados = set(
            self.limpar_termo(os.path.splitext(os.path.basename(arquivo))[0].split('-')[0])
            for arquivo in encontrados
        )

        self.arquivos_nao_encontrados = [
            t for t in termos_busca if t not in termos_encontrados
        ]

        if self.arquivos_nao_encontrados:
            self.status_var.set(f"Concluído! {len(encontrados)} imagens. Não encontradas: {len(self.arquivos_nao_encontrados)}")
        else:
            self.status_var.set(f"Sucesso! Todas {len(encontrados)} imagens processadas.")

        self.rodando = False
        if encontrados:
            os.startfile(DOWNLOAD_PATH)

    def criar_botao_round(self, texto, comando, x, y, largura=220, altura=40):
        canvas = Canvas(self, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
        canvas.place(x=x, y=y)
        canvas.create_oval(0, 0, 20, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        canvas.create_oval(largura-20, 0, largura, altura, fill=COR_BOTAO, outline=COR_BOTAO)
        canvas.create_rectangle(10, 0, largura-10, altura, fill=COR_BOTAO, outline=COR_BOTAO)

        botao = Button(self, text=texto, command=comando, bg=COR_BOTAO, fg=COR_BOTAO_TEXTO,
                       font=FONTE_PADRAO, relief="flat", bd=0, activebackground=COR_BOTAO)
        botao.place(x=x + 10, y=y + 8)
        return botao

    def animar_bolinha(self):
        if not self.rodando:
            self.bolinha_canvas.delete("all")
            return

        if self.indo_para_direita:
            self.bolinha_pos += 4
            if self.bolinha_pos > 80:
                self.indo_para_direita = False
        else:
            self.bolinha_pos -= 4
            if self.bolinha_pos < 0:
                self.indo_para_direita = True

        self.bolinha_canvas.delete("all")
        self.bolinha_canvas.create_oval(10 + self.bolinha_pos, 10, 22 + self.bolinha_pos, 22,
                                        fill=COR_BOTAO, outline=COR_BOTAO)
        self.bolinha_canvas.after(30, self.animar_bolinha)

    def iniciar_processamento(self):
        if not self.entry_excel.get():
            self.status_var.set("Selecione um arquivo Excel primeiro!")
            return

        if not self.coluna_selecionada.get():
            self.status_var.set("Selecione uma coluna válida!")
            return

        if self.rodando:
            return

        self.rodando = True
        threading.Thread(target=self.animar_bolinha, daemon=True).start()
        threading.Thread(target=self.gerar_zip, daemon=True).start()

    def limpar_campos(self):
        self.entry_excel.delete(0, "end")
        self.coluna_selecionada.set("")
        self.status_var.set("")
        self.arquivos_nao_encontrados = []
        self.df_planilha = None

        if self.option_menu:
            self.option_menu.destroy()
            self.option_menu = None

    def criar_logo(self):
        try:
            image_path = resource_path("assets/altenburg.png")
            img_original = Image.open(image_path).resize((210, 210), Image.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img_original)  # Armazena como atributo da classe
            logo_label = Label(self, image=self.logo_img, bg=COR_FUNDO_APP, borderwidth=0)
            logo_label.image = self.logo_img
            logo_label.place(x=10, y=18)
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")

    def criar_interface(self):
        self.criar_logo()
        
        Label(self, text="| CELLY - Buscador de Imagens", bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE,
              font=("Segoe UI Semibold", 13)).place(x=210, y=101)

        Label(self, text="📄 Caminho do arquivo Excel:", bg=COR_FUNDO_APP,
              fg=COR_TEXTO_FORTE, font=FONTE_PADRAO, anchor="w").place(x=20, y=160)

        self.entry_excel = Entry(self, font=FONTE_PADRAO, relief="flat", bg="white")
        self.entry_excel.place(x=20, y=190, width=300)

        self.criar_botao_round(
            "     📂  Selecionar Arquivo",
            lambda: [
                self.entry_excel.delete(0, "end"),
                self.entry_excel.insert(0, filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])),
                self.carregar_colunas()
            ],
            x=315, y=230
        )

        botao_lixeira = Button(self, text="🗑️", command=self.limpar_campos,
                               bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO,
                               relief="flat", bd=0, activebackground=COR_FUNDO_APP,
                               activeforeground="white", highlightthickness=0)
        botao_lixeira.place(x=534, y=258, width=50, height=40)

        self.criar_botao_round("     🧾  Gerar ZIP com imagens", self.iniciar_processamento, x=315, y=280)

        self.bolinha_canvas = Canvas(self, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
        self.bolinha_canvas.place(x=315, y=320)

        Label(self, textvariable=self.status_var, bg=COR_FUNDO_APP, fg=COR_TEXTO_FORTE,
              font=FONTE_PADRAO).place(x=20, y=340)

        Label(self, text="© DEVELOPED BY AMANDA PRADO", bg=COR_FUNDO_APP,
              fg=COR_TEXTO_FORTE, font=("Segoe UI", 7)).place(x=319, y=380, anchor="n")

        botao_casinha = Button(self, text="🏠", command=lambda: self.controller.show_frame("MenuFrame"),
                               bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO,
                               relief="flat", bd=0, activebackground=COR_FUNDO_APP,
                               activeforeground="white", highlightthickness=0)
        botao_casinha.place(x=-8, y=1, width=50, height=40)