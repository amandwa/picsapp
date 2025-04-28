import os
import pandas as pd
import webbrowser
from email.message import EmailMessage
from zipfile import ZipFile
from tkinter import Tk, filedialog, Label, Entry, Button
import re

# Pastas onde buscar (toda a estrutura, com subpastas)
PASTAS_BUSCA = [r"L:\\IMAGENS GERAL", r"E:\\"]

# Função para normalizar nomes (remove pontuação, espaços, traços, etc.)
def normalizar_nome(nome):
    return re.sub(r'[^a-zA-Z0-9]', '', nome)

# Função para gerar o e-mail com o ZIP anexado
def gerar_email(completo_arquivo_zip, email_cliente):
    corpo_email = """\
Olá, estimado cliente!

Segue em anexo as imagens relacionadas aos produtos Altenburg que você adquiriu/comprou.

Atenciosamente,
"""

    msg = EmailMessage()
    msg["Subject"] = "Fotos dos produtos adquiridos"
    msg["From"] = "seu_email@outlook.com"
    msg["To"] = email_cliente
    msg.set_content(corpo_email)

    with open(completo_arquivo_zip, "rb") as f:
        zip_data = f.read()
        zip_name = os.path.basename(completo_arquivo_zip)
        msg.add_attachment(zip_data, maintype="application", subtype="zip", filename=zip_name)

    eml_path = "email_com_arquivo.eml"
    with open(eml_path, "wb") as f:
        f.write(bytes(msg))

    webbrowser.open(f'file://{os.path.abspath(eml_path)}')

# Função principal
def processar_fotos():
    excel_path = entry_excel.get()
    email_cliente = entry_email.get()

    df = pd.read_excel(excel_path, dtype=str)

    # Extrair termos das colunas A (0), C (2) e D (3)
    colunas_desejadas = [0, 2, 3]
    termos_busca = []
    for col in colunas_desejadas:
        if col < len(df.columns):
            termos_busca += df.iloc[:, col].dropna().astype(str).tolist()

    # Normalizar termos
    termos_normalizados = [normalizar_nome(t) for t in termos_busca]

    # Criar ZIP
    zip_nome = "fotos_cliente.zip"
    encontrados = set()

    with ZipFile(zip_nome, 'w') as zipf:
        for pasta in PASTAS_BUSCA:
            for raiz, _, arquivos in os.walk(pasta):
                for arquivo in arquivos:
                    nome_arquivo_normalizado = normalizar_nome(arquivo)
                    for termo in termos_normalizados:
                        if termo in nome_arquivo_normalizado and arquivo not in encontrados:
                            caminho_arquivo = os.path.join(raiz, arquivo)
                            zipf.write(caminho_arquivo, os.path.basename(caminho_arquivo))
                            encontrados.add(arquivo)
                            break

    if not encontrados:
        print("Nenhuma imagem encontrada com os termos fornecidos.")
    else:
        print(f"{len(encontrados)} imagem(ns) adicionada(s) ao ZIP.")

    gerar_email(zip_nome, email_cliente)

    os.remove(zip_nome)
    print("Processo finalizado.")

# Selecionar Excel
def selecionar_excel():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_excel.delete(0, "end")
    entry_excel.insert(0, file)

# Tkinter UI
root = Tk()
root.title("Envio de Fotos para Cliente")

Label(root, text="E-mail do cliente:").pack(pady=5)
entry_email = Entry(root, width=50)
entry_email.pack(pady=5)

Label(root, text="Arquivo Excel (contendo termos nas colunas A, C e D):").pack(pady=5)
entry_excel = Entry(root, width=50)
entry_excel.pack(pady=5)
Button(root, text="Selecionar Excel", command=selecionar_excel).pack(pady=5)

Button(root, text="Gerar E-mail", command=processar_fotos).pack(pady=20)

root.mainloop()
