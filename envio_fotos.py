import os
import pandas as pd
import re
from tkinter import Tk, Canvas, Label, Button, filedialog, Entry
from zipfile import ZipFile
import threading

# Função para selecionar o Excel
def selecionar_excel():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_excel.delete(0, "end")
    entry_excel.insert(0, file)
    msg_arquivo_selecionado.config(text="Arquivo selecionado!")

# Função para selecionar a pasta com imagens
def selecionar_pasta():
    pasta = filedialog.askdirectory()
    entry_pasta.delete(0, "end")
    entry_pasta.insert(0, pasta)
    msg_pasta_selecionada.config(text="Pasta selecionada!")

# Limpa o termo: remove símbolos e deixa apenas letras/números
def limpar_termo(termo):
    return re.sub(r'[^a-zA-Z0-9]', '', termo)

# Função principal para gerar o ZIP
def gerar_zip():
    caminho_excel = entry_excel.get()
    caminho_pasta = entry_pasta.get()

    if not os.path.exists(caminho_excel):
        msg_status.config(text="Excel não encontrado.", fg="red")
        return
    if not os.path.exists(caminho_pasta):
        msg_status.config(text="Pasta inválida.", fg="red")
        return

    try:
        df = pd.read_excel(caminho_excel, dtype=str)
    except Exception as e:
        msg_status.config(text=f"Erro ao ler Excel: {e}", fg="red")
        return

    # Coleta todos os termos da planilha
    termos = set()
    for _, row in df.iterrows():
        for val in row:
            if pd.notna(val):
                termo = limpar_termo(str(val).strip().lower())
                if termo:
                    termos.add(termo)

    caminho_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    zip_nome = os.path.join(caminho_downloads, "fotos_encontradas.zip")

    encontrados = 0
    with ZipFile(zip_nome, 'w') as zipf:
        for root, dirs, files in os.walk(caminho_pasta):
            for file in files:
                nome_base = os.path.splitext(file)[0].lower()
                nome_limpo = limpar_termo(nome_base)

                for termo in termos:
                    if termo in nome_limpo:
                        caminho_completo = os.path.join(root, file)
                        zipf.write(caminho_completo, file)
                        encontrados += 1
                        break  # já encontrou correspondência

    if encontrados:
        msg_status.config(text=f"{encontrados} imagens encontradas e salvas em {zip_nome}", fg="green")
    else:
        msg_status.config(text="Nenhuma imagem encontrada.", fg="orange")

# Roda o processo em uma thread separada (para evitar travar a interface)
def rodar_em_thread():
    threading.Thread(target=gerar_zip).start()

# Interface gráfica
root = Tk()
root.title("Buscador de Imagens por Excel")
root.geometry("600x300")
root.configure(bg="#f4e8d4")
root.resizable(False, False)

# Interface Excel
Label(root, text="Selecione o arquivo Excel:", bg="#f4e8d4").place(x=20, y=20)
entry_excel = Entry(root, width=50)
entry_excel.place(x=20, y=50)
Button(root, text="Selecionar Excel", command=selecionar_excel).place(x=450, y=47)
msg_arquivo_selecionado = Label(root, text="", bg="#f4e8d4", fg="green")
msg_arquivo_selecionado.place(x=20, y=75)

# Interface pasta
Label(root, text="Selecione a pasta com imagens:", bg="#f4e8d4").place(x=20, y=110)
entry_pasta = Entry(root, width=50)
entry_pasta.place(x=20, y=140)
Button(root, text="Selecionar Pasta", command=selecionar_pasta).place(x=450, y=137)
msg_pasta_selecionada = Label(root, text="", bg="#f4e8d4", fg="green")
msg_pasta_selecionada.place(x=20, y=165)

# Botão principal
Button(root, text="Gerar ZIP com imagens", command=rodar_em_thread, bg="#2e2e2e", fg="white").place(x=200, y=200)

# Mensagem de status
msg_status = Label(root, text="", bg="#f4e8d4", fg="black", font=("Segoe UI", 10))
msg_status.place(x=20, y=240)

root.mainloop()
