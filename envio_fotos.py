import os
import pandas as pd
import smtplib
from email.message import EmailMessage
from zipfile import ZipFile
from tkinter import Tk, filedialog, Label, Entry, Button

#  Função para enviar o e-mail 
def enviar_email(completo_arquivo_zip, email_cliente):
    msg = EmailMessage()
    msg['Subject'] = 'Fotos dos produtos adquiridos'
    msg['From'] = 'seu_email@outlook.com'
    msg['To'] = email_cliente
    msg.set_content("Olá, segue em anexo as fotos dos produtos adquiridos.")

    #  Anexar o arquivo zip com as fotos
    with open(completo_arquivo_zip, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='zip', filename=completo_arquivo_zip)

    #  Enviar e-mail via SMTP do Outlook
    try:
        with smtplib.SMTP_SSL('smtp-mail.outlook.com', 465) as smtp:
            smtp.login('seu_email@outlook.com', 'sua_senha')
            smtp.send_message(msg)
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

#  Função para gerar e processar o zip com as fotos 
def processar_fotos():
    #  Obter os dados inseridos
    pasta_fotos = entry_pasta.get()
    email_cliente = entry_email.get()
    arquivo_excel = entry_excel.get()


    #  Ler o arquivo Excel para extrair os nome das fotos 
    df = pd.read_excel(arquivo_excel)
    nomes_fotos = df.iloc[:, 0].astype(str).tolist()

    #  Pasta temporária para armazenar as fotos
    pasta_temp = "fotos_temp"
    os.makedirs(pasta_temp, exist_ok=True)

    # Adicionar fotos no arquivo zip
    zip_nome = "fotos_cliente.zip"
    with ZipFile(zip_nome, 'w') as zipf:
        for nome_foto in nomes_fotos:
            caminho_foto = os.path.join(pasta_fotos, nome_foto)
            if os.path.exists(caminho_foto):
                zipf.write(caminho_foto, os.path.basename(caminho_foto))

    # Enviar o arquivo zip por e-mail
    enviar_email(zip_nome, email_cliente)

    # Limpeza após o processo
    os.remove(zip_nome)
    print("Processo completo!")

# Função para selecionar a pasta de fotos
def selecionar_pasta():
    pasta = filedialog.askdirectory()
    entry_pasta.delete(0, "end")
    entry_pasta.insert(0, pasta)

#  Função para selecionar o arquivo Excel
def selecionar_excel():
    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_excel.delete(0, "end")
    entry_excel.insert(0, excel_file)

#  Interface de teste, b4 the front part 
root = Tk()
root.title("Envio de Fotos para Cliente")

# Campos de entrada
Label(root, text="Caminho da pasta com as fotos:").pack(pady=5)
entry_pasta = Entry(root, width=50)
entry_pasta.pack(pady=5)
Button(root, text="Selecionar Pasta", command=selecionar_pasta).pack(pady=5)

Label(root, text="E-mail do cliente:").pack(pady=5)
entry_email = Entry(root, width=50)
entry_email.pack(pady=5)

Label(root, text="Arquivo Excel (com nomes das fotos):").pack(pady=5)
entry_excel = Entry(root, width=50)
entry_excel.pack(pady=5)
Button(root, text="Selecionar Excel", command=selecionar_excel).pack(pady=5)

# Botão para processar e enviar
Button(root, text="Enviar Fotos", command=processar_fotos).pack(pady=20)

# Iniciar a interface
root.mainloop()
 