import pandas as pd
import os
from tkinter import Tk, Label, Entry, Canvas, Button, filedialog, messagebox

# Paleta elegante com fundo roxo escuro
COR_FUNDO_APP = "#252440"
COR_TEXTO = "#FFFFFF"
COR_BOTAO = "#B497D6"
COR_BOTAO_TEXTO = "#252440"
FONTE_PADRAO = ("Segoe UI", 10)  # Fonte moderna e nativa do Windows

def criar_botao_round(master, texto, comando, x, y, largura=180, altura=40, cor=COR_BOTAO, cor_texto=COR_BOTAO_TEXTO):
    canvas = Canvas(master, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=cor, outline=cor)
    canvas.create_oval(largura-20, 0, largura, altura, fill=cor, outline=cor)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=cor, outline=cor)
    botao = Button(master, text=texto, command=comando, bg=cor, fg=cor_texto, font=FONTE_PADRAO, relief="flat", bd=0, activebackground=cor)
    botao.place(x=x + 5, y=y + 8)
    return botao

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho:
        entrada_arquivo.delete(0, "end")
        entrada_arquivo.insert(0, caminho)

def processar_arquivo():
    caminho_arquivo = entrada_arquivo.get()
    separador = entrada_separador.get()

    if not caminho_arquivo or not separador:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo e informe o separador.")
        return

    try:
        df = pd.read_excel(caminho_arquivo, dtype=str)

        novas_linhas = []

        # Processa cada linha do DataFrame
        for _, row in df.iterrows():
            # Verifica e divide os valores nas colunas F e G com o separador fornecido
            itens = str(row.iloc[5]).split(separador)  # Coluna F
            quantidades = str(row.iloc[6]).split(separador)  # Coluna G

            # Limpa espaços extras e mantém apenas valores não vazios
            itens = [item.strip() for item in itens if item.strip()]
            quantidades = [qtd.strip() for qtd in quantidades if qtd.strip()]

            # Cria novas linhas combinando os itens e quantidades
            for i in range(len(itens)):
                nova_linha = row.copy()
                nova_linha.iloc[5] = itens[i]
                nova_linha.iloc[6] = quantidades[i] if i < len(quantidades) else ""
                novas_linhas.append(nova_linha)

        df_novo = pd.DataFrame(novas_linhas)

        # Caminho para salvar o arquivo modificado
        pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        caminho_saida = os.path.join(pasta_downloads, "alteracao_ticket.xlsx")
        df_novo.to_excel(caminho_saida, index=False)

        # Mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho_saida}")

    except Exception as e:
        messagebox.showerror("Erro ao processar", str(e))

# Interface gráfica
app = Tk()
app.title("Separador de Colunas")
app.geometry("540x250")
app.configure(bg=COR_FUNDO_APP)
app.resizable(False, False)

Label(app, text="Arquivo Excel:", bg=COR_FUNDO_APP, fg=COR_TEXTO, font=FONTE_PADRAO).place(x=30, y=30)
entrada_arquivo = Entry(app, width=50, font=FONTE_PADRAO)
entrada_arquivo.place(x=150, y=32)
criar_botao_round(app, "Selecionar", selecionar_arquivo, x=400, y=27, largura=100, altura=35)

Label(app, text="Separador (ex: ;)", bg=COR_FUNDO_APP, fg=COR_TEXTO, font=FONTE_PADRAO).place(x=30, y=90)
entrada_separador = Entry(app, width=10, font=FONTE_PADRAO)
entrada_separador.place(x=180, y=92)

criar_botao_round(app, "Processar", processar_arquivo, x=180, y=150, largura=180, altura=40)

app.mainloop()
