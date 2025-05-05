import os
import pandas as pd
import re
from tkinter import Tk, Canvas, Label, Button, filedialog, Entry, StringVar, OptionMenu
from zipfile import ZipFile
import threading
import shutil
from PIL import Image, ImageTk
import sys
import subprocess

# ========== CONSTANTES ==========
COR_TEXTO_FORTE = "#FFFFFF"
COR_FUNDO_APP = "#252440"
COR_BOTAO = "#D0E7FF"
COR_BOTAO_TEXTO = "#252440"
COR_BARRA_TITULO = "#2D2D57"
FONTE_PADRAO = ("Segoe UI", 10)
DOWNLOAD_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
CAMINHO_PASTA_IMAGENS = "E:\\"

# ========== CONFIGURAÇÃO DA JANELA ==========
root = Tk()
root.title("Buscador de Imagens por Excel")
root.geometry("650x400")
root.configure(bg=COR_FUNDO_APP)
root.resizable(False, False)
root.overrideredirect(True)

# ========== VARIÁVEIS GLOBAIS ==========
rodando = False
arquivos_nao_encontrados = []
df_planilha = None
option_menu = None
bolinha_pos = 0
indo_para_direita = True
coluna_selecionada = StringVar(root)
status_var = StringVar(root)

# ========== FUNÇÕES PRINCIPAIS ==========
def limpar_termo(termo):
    """Remove caracteres especiais e formata para comparação."""
    termo = str(termo).split('-')[0]  # Pega apenas a parte antes do primeiro hífen
    return re.sub(r'[^0-9]', '', termo)  # Mantém apenas números

def carregar_colunas():
    global df_planilha, option_menu
    
    caminho_excel = entry_excel.get()
    if not os.path.exists(caminho_excel):
        status_var.set("Arquivo não encontrado!")
        return
    
    try:
        df_planilha = pd.read_excel(caminho_excel, dtype=str)
        colunas = list(df_planilha.columns)
        coluna_selecionada.set(colunas[0] if colunas else "")
        
        if option_menu:
            option_menu.destroy()
            
        option_menu = OptionMenu(root, coluna_selecionada, *colunas)
        option_menu.config(bg="white", font=FONTE_PADRAO, highlightthickness=0)
        option_menu.place(x=20, y=250)
        
        Label(root, text="Selecione a coluna 'Referência + Derivação':", 
              bg=COR_FUNDO_APP, fg="white", font=FONTE_PADRAO).place(x=20, y=220)
        
        status_var.set("Planilha carregada com sucesso!")
    except Exception as e:
        status_var.set(f"Erro ao carregar planilha: {str(e)}")

def gerar_zip():
    global rodando, arquivos_nao_encontrados
    
    # Verificações iniciais
    caminho_excel = entry_excel.get()
    if not os.path.exists(caminho_excel):
        status_var.set("Arquivo Excel não encontrado!")
        rodando = False
        return
    
    if not os.path.exists(CAMINHO_PASTA_IMAGENS):
        status_var.set(f"Pasta de imagens não encontrada em: {CAMINHO_PASTA_IMAGENS}")
        rodando = False
        return

    try:
        df = pd.read_excel(caminho_excel, dtype=str)
    except Exception as e:
        status_var.set(f"Erro ao ler planilha: {str(e)}")
        rodando = False
        return

    coluna = coluna_selecionada.get()
    if coluna not in df.columns:
        status_var.set("Coluna selecionada não existe na planilha!")
        rodando = False
        return

    # Criar dicionário de termos de busca
    termos_busca = {}
    for _, row in df.iterrows():
        valor = str(row[coluna]) if coluna in row and pd.notna(row[coluna]) else ""
        if valor:
            termo = limpar_termo(valor)
            if termo:  # Só adiciona se não for vazio após limpeza
                termos_busca[termo] = row[coluna]  # Usa o valor original como nome do arquivo

    # Buscar imagens
    encontrados = []
    contador_nomes = {}
    
    for root_dir, _, files in os.walk(CAMINHO_PASTA_IMAGENS):
        for file in files:
            if not rodando:
                break
                
            nome_arquivo = os.path.splitext(file)[0]
            nome_limpo = limpar_termo(nome_arquivo)
            
            # Verifica cada termo de busca
            for termo_busca, nome_final_base in termos_busca.items():
                if termo_busca in nome_limpo:
                    ext = os.path.splitext(file)[1].lower()
                    
                    # Remove a parte após o hífen do nome final
                    nome_base_sem_derivacao = nome_final_base.split('-')[0]
                    
                    # Contador para evitar sobrescrita
                    contador = contador_nomes.get(nome_base_sem_derivacao, 0)
                    nome_final = f"{nome_base_sem_derivacao}_{contador}{ext}" if contador > 0 else f"{nome_base_sem_derivacao}{ext}"
                    contador_nomes[nome_base_sem_derivacao] = contador + 1
                    
                    origem = os.path.join(root_dir, file)
                    destino = os.path.join(DOWNLOAD_PATH, nome_final)
                    
                    try:
                        shutil.copy2(origem, destino)
                        encontrados.append(destino)
                        status_var.set(f"Encontrado: {nome_final}")
                        root.update()
                    except Exception as e:
                        print(f"Erro ao copiar {origem}: {str(e)}")

    # Criar ZIP
    if encontrados:
        zip_path = os.path.join(DOWNLOAD_PATH, "catalogo_imagens.zip")
        try:
            with ZipFile(zip_path, 'w') as zipf:
                for arquivo in encontrados:
                    if os.path.exists(arquivo):
                        zipf.write(arquivo, os.path.basename(arquivo))
                        os.remove(arquivo)
                        status_var.set(f"Adicionando ao ZIP: {os.path.basename(arquivo)}")
                        root.update()
        except Exception as e:
            status_var.set(f"Erro ao criar ZIP: {str(e)}")
            rodando = False
            return
    
    # Resultado final
    arquivos_nao_encontrados = [t for t in termos_busca if t not in [limpar_termo(f.split('-')[0]) for f in encontrados]]
    
    if arquivos_nao_encontrados:
        status_var.set(f"Concluído! {len(encontrados)} imagens. Não encontradas: {len(arquivos_nao_encontrados)}")
    else:
        status_var.set(f"Sucesso! Todas {len(encontrados)} imagens processadas.")
    
    rodando = False
    if encontrados:
        os.startfile(DOWNLOAD_PATH)

def criar_botao_round(master, texto, comando, x, y, largura=220, altura=40):
    canvas = Canvas(master, width=largura, height=altura, bg=COR_FUNDO_APP, highlightthickness=0)
    canvas.place(x=x, y=y)
    canvas.create_oval(0, 0, 20, altura, fill=COR_BOTAO, outline=COR_BOTAO)
    canvas.create_oval(largura-20, 0, largura, altura, fill=COR_BOTAO, outline=COR_BOTAO)
    canvas.create_rectangle(10, 0, largura-10, altura, fill=COR_BOTAO, outline=COR_BOTAO)
    
    botao = Button(master, text=texto, command=comando, bg=COR_BOTAO, fg=COR_BOTAO_TEXTO, 
                   font=FONTE_PADRAO, relief="flat", bd=0, activebackground=COR_BOTAO)
    botao.place(x=x + 10, y=y + 8)
    return botao

def animar_bolinha():
    global bolinha_pos, indo_para_direita
    
    if not rodando:
        bolinha_canvas.delete("all")
        return
    
    if indo_para_direita:
        bolinha_pos += 4
        if bolinha_pos > 80:
            indo_para_direita = False
    else:
        bolinha_pos -= 4
        if bolinha_pos < 0:
            indo_para_direita = True
    
    bolinha_canvas.delete("all")
    bolinha_canvas.create_oval(10 + bolinha_pos, 10, 22 + bolinha_pos, 22, fill=COR_BOTAO, outline=COR_BOTAO)
    bolinha_canvas.after(30, animar_bolinha)

def iniciar_processamento():
    global rodando
    
    if not entry_excel.get():
        status_var.set("Selecione um arquivo Excel primeiro!")
        return
    
    if not coluna_selecionada.get():
        status_var.set("Selecione uma coluna válida!")
        return
    
    if rodando:
        return
    
    rodando = True
    threading.Thread(target=animar_bolinha, daemon=True).start()
    threading.Thread(target=gerar_zip, daemon=True).start()

def limpar_campos():
    global arquivos_nao_encontrados, df_planilha, option_menu
    
    entry_excel.delete(0, "end")
    coluna_selecionada.set("")
    status_var.set("")
    arquivos_nao_encontrados = []
    df_planilha = None
    
    if option_menu:
        option_menu.destroy()
        option_menu = None

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
    text="| CELLY - Buscador de Imagens",
    bg=COR_FUNDO_APP,
    fg=COR_TEXTO_FORTE,
    font=("Segoe UI Semibold", 13)
)
texto_logo.place(x=210, y=126)


# ========== COMPONENTES DA INTERFACE ==========
# Barra superior
barra_superior = Canvas(root, width=650, height=35, bg=COR_BARRA_TITULO, highlightthickness=0)
barra_superior.place(x=0, y=0)

def mover_janela(event):
    x = event.x_root - root.x
    y = event.y_root - root.y
    root.geometry(f"+{x}+{y}")

def iniciar_mover(event):
    root.x = event.x
    root.y = event.y

barra_superior.bind("<Button-1>", iniciar_mover)
barra_superior.bind("<B1-Motion>", mover_janela)

botao_fechar = Button(barra_superior, text="X", command=root.quit, bg=COR_BARRA_TITULO, 
                      fg=COR_TEXTO_FORTE, font=("Arial", 12, "bold"), relief="flat", bd=0)
botao_fechar.place(x=615, y=5)

def voltar_menu(nome_script="iniciar.py"):
    """
    Fecha o app atual e abre o menu principal (ou outro script especificado).
    """
    caminho_script = os.path.join(os.path.dirname(__file__), nome_script)
    subprocess.Popen([sys.executable, caminho_script])
    root.destroy()

# Componentes da interface
Label(
    root,
    text="📄 Caminho do arquivo Excel:",
    bg=COR_FUNDO_APP,
    fg=COR_TEXTO_FORTE,
    font=FONTE_PADRAO,
    anchor="w"  # alinha o texto à esquerda
).place(x=20, y=160) 


entry_excel = Entry(root, font=FONTE_PADRAO, relief="flat", bg="white")
entry_excel.place(x=20, y=190, width=300)

criar_botao_round(
    root, "     📂  Selecionar Arquivo",
    lambda: [
        entry_excel.delete(0, "end"),
        entry_excel.insert(0, filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])),
        carregar_colunas()
    ],
    x=350, y=230
)

botao_lixeira = Button(
    root,
    text="🗑️",
    command=limpar_campos,
    bg=COR_FUNDO_APP,           # mesmo fundo do app
    fg="white",
    font=FONTE_PADRAO,
    activebackground=COR_FUNDO_APP,  # fundo não muda ao clicar
    activeforeground="white",
    relief="flat",
    bd=0,
    cursor="hand2",
    highlightthickness=0,
    takefocus=0
)
botao_lixeira.place(x=570, y=258, width=50, height=40)


criar_botao_round(root, "     🧾  Gerar ZIP com imagens", iniciar_processamento, x=350, y=280)

bolinha_canvas = Canvas(root, width=110, height=32, bg=COR_FUNDO_APP, highlightthickness=0)
bolinha_canvas.place(x=350, y=320)

status_label = Label(root, textvariable=status_var, bg=COR_FUNDO_APP, 
                    fg=COR_TEXTO_FORTE, font=FONTE_PADRAO)
status_label.place(x=20, y=340)

Label(root, text="© DEVELOPED BY AMANDA PRADO", bg=COR_FUNDO_APP, 
      fg=COR_TEXTO_FORTE, font=("Segoe UI", 7)).place(x=319, y=380, anchor="n")

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