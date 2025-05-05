import tkinter as tk
from frames.menu import MenuFrame
from frames.imagens import ImagensFrame
from frames.colunas import ColunasFrame
import os
import sys

# Função para obter caminho absoluto (compatível com PyInstaller)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Diretório temporário criado pelo PyInstaller
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Janela principal
        self.title("CELLY")
        self.geometry("600x400")
        self.configure(bg="#f0f0f0")
        self.resizable(False, False)

        # Ícone da aplicação
        icon_path = resource_path("assets/icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception as e:
                print(f"Erro ao definir ícone: {e}")

        # Container principal
        container = tk.Frame(self, width=600, height=400, bg="#f0f0f0")
        container.pack(fill="both", expand=True)

        # Dicionário de frames
        self.frames = {}

        # Inicializa todos os frames e os empilha
        for FrameClass in (MenuFrame, ImagensFrame, ColunasFrame):
            page_name = FrameClass.__name__
            frame = FrameClass(parent=container, controller=self)
            frame.place(x=0, y=0, width=600, height=400)
            self.frames[page_name] = frame

        # Exibe a tela inicial
        self.show_frame("MenuFrame")

    def show_frame(self, page_name):
        """Exibe o frame com base no nome da classe."""
        frame = self.frames.get(page_name)
        if frame:
            frame.tkraise()
        else:
            print(f"Frame '{page_name}' não encontrado.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
