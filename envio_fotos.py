import os
import pandas as pd
import webbrowser
from email.message import EmailMessage
from zipfile import ZipFile
from tkinter import Tk, filedialog, Label, Entry, Button

# Caminhos fixos
CAMINHO_EXCEL = "C:\\caminho\\fixo\\arquivo.xlsx"  # ajuste conforme seu sistema
PASTAS_BUSCA = [r"L:\\IMAGENS GERAL", r"E:\\"]  # caminhos onde procurar imagens
EMAIL_DESTINATARIO = "cliente@exemplo.com"  # e-mail fixo ou fictício
PASTA_TEMP = "fotos_temp"

# Função para gerar e abrir o .eml com o anexo ZIP
def gerar_email(completo_arquivo_zip):
    corpo_email = """\
Olá, estimado cliente!

Segue em anexo as imagens relacionadas aos produtos Altenburg que você adquiriu/comprou.

Atenciosamente,
"""

    # Criar estrutura do e-mail
    msg = EmailMessage()
    msg["Subject"] = "Fotos dos produtos adquiridos"
    msg["From"] = "seu_email@outlook.com"
    msg["To"] = EMAIL_DESTINATARIO
    msg.set_content(corpo_email)

    with open(completo_arquivo_zip, "rb") as f:
        zip_data = f.read()
        zip_name = os.path.basename(completo_arquivo_zip)
        msg.add_attachment(zip_data, maintype="application", subtype="zip", filename=zip_name)

    # Salvar o e-mail como .eml
    eml_path = "email_com_arquivo.eml"
    with open(eml_path, "wb") as f:
        f.write(bytes(msg))

    # Abrir no cliente de e-mail padrão
    webbrowser.open(f'file://{os.path.abspath(eml_path)}')

# Função principal
def processar_fotos():
    # Ler Excel
    df = pd.read_excel(CAMINHO_EXCEL)
    nomes_fotos = df['Referência + Derivação'].astype(str).tolist()

    # Criar diretório temporário se não existir
    os.makedirs(PASTA_TEMP, exist_ok=True)

    # Criar ZIP
    zip_nome = "fotos_cliente.zip"
    with ZipFile(zip_nome, 'w') as zipf:
        for nome_foto in nomes_fotos:
            foto_encontrada = False
            for pasta in PASTAS_BUSCA:
                caminho_foto = os.path.join(pasta, nome_foto)
                if os.path.exists(caminho_foto):
                    zipf.write(caminho_foto, os.path.basename(caminho_foto))
                    foto_encontrada = True
                    break
            if not foto_encontrada:
                print(f"Imagem {nome_foto} não foi encontrada.")

    # Gerar e-mail com anexo
    gerar_email(zip_nome)

    # Limpeza
    os.remove(zip_nome)
    print("Processo finalizado!")

# Executar diretamente
if __name__ == "__main__":
    processar_fotos()