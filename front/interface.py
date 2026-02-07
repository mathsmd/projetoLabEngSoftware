from tkinter import Tk, Label, ttk
from PIL import Image, ImageTk

class Interface:
    def __init__(self):
        self.root = Tk()
        self.indice = 0

        self.imagens = [
            ImageTk.PhotoImage(file="front/images/cabecinha1.png"),
            ImageTk.PhotoImage(file="front/images/cabecinha2.png")
        ]

        self.tela()
        self.frames_da_tela()
        self.alternar()

    def tela(self):
        self.root.title('Cabecinha LEDS')
        self.root.attributes('-topmost', True, '-toolwindow', True, '-alpha', 1.0)
        self.root.geometry('240x160+0+0')
        self.root.resizable(False, False)

    def frames_da_tela(self):
        self.sprite = Label(self.root, image=self.imagens[self.indice])
        self.sprite.pack()

        self.label = Label(self.root, text='Iniciando...')
        self.label.pack()

        self.pgrbar = ttk.Progressbar(self.root, length=200)
        self.pgrbar.pack()

    def upgrade_bar(self, texto, progresso):
        self.label.config(text=texto)
        self.pgrbar['value'] += progresso

    def alternar(self):
        self.indice = (self.indice + 1) % len(self.imagens)
        self.sprite.config(image=self.imagens[self.indice])
        self.root.after(1000, self.alternar)

    def start(self):
        self.root.mainloop()


