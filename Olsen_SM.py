from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import pandas as pd
import pywhatkit

# Getting paths
icon = os.path.join('.', 'images', 'icon.ico')
path_img = None

# functions

def selecionar_planilha():
    global path_pl
    diretorio_planilha = filedialog.askopenfile(
    title="Selecionar Diretório",
    initialdir=".",
    filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if diretorio_planilha:
        path_pl = diretorio_planilha.name
        diretoriop_texto.insert(END, path_pl)

def selecionar_imagem():
    global path_img
    diretorio_img = filedialog.askopenfile(
    title="Selecionar Diretório",
    initialdir=".",
    filetypes=[("Imagens", "*.JPEG *.JPG *.PNG")]
    )
    if diretorio_img:
        path_img = diretorio_img.name
        diretorioimg_texto.insert(END, path_img)
    
def enviar_mensagem(path_img=None):
    # Função para exibir a caixa de confirmação
    resposta = messagebox.askquestion("Confirmar envio", "Tem certeza que deseja enviar a mensagem?")
    if resposta == "no":
        messagebox.showwarning('Aviso', 'Envio cancelado.')
    else:
        df = pd.read_excel(path_pl, dtype=str)
        for index, row in df.iterrows():
            phone_number = row['Número']
            name = row['Nome']
            message = entry_msg.get("1.0", END)
            message = message.replace('<nome>', name)
            if path_img is None:
                pywhatkit.sendwhatmsg_instantly(phone_number, message=message, wait_time=10, tab_close=True)
            else:
                pywhatkit.sendwhats_image(phone_number, path_img, caption=message, wait_time=12, tab_close=True)

# Creating the app

root = Tk()
root.resizable(width=False, height=False)
root.title('OlsenK - Sending Messages')
root.iconbitmap(icon) # icone
root.geometry('500x250+250+250') # definindo tamanho da janela

################# botao selecionar planilha ###################################################

## botão selecionar planilha
btn_selecionar = Button(root, text="Selecionar Planilha", command=selecionar_planilha, width=15)
btn_selecionar.place(x=10, y=10)

## campo demonstrando dir da planilha
diretoriop_texto = Text(root, width=30, height=0)
diretoriop_texto.place(x=130, y=13)

################# botao selecionar imagem ###################################################
## botão selecionar imagem
btn_selecionar_img = Button(root, text="Selecionar Imagem", command=selecionar_imagem, width=15)
btn_selecionar_img.place(x=10, y=40)

## campo demonstrando dir da imagem
diretorioimg_texto = Text(root, width=30, height=0)
diretorioimg_texto.place(x=130, y=43)

################# campo mensagem ###################################################
# alerta campo mensagem
alerta = Label(text=f"Campo Enviar Mensagem - Colocar a variável como '<nome>'")
alerta.place(x=10, y=80)

# campo mensagem
my_font=("Segoe UI", 10)
entry_msg = Text(root, width=52, height=5, font=my_font)
entry_msg.place(x=10, y=100)

################# botao enviar imagem ###################################################

## botao enviar mensagens
btn_envio = Button(root, text="Enviar mensagens", command=lambda: enviar_mensagem(path_img), width=15)
btn_envio.place(x=370, y=210)

# obrigatório
root.mainloop()