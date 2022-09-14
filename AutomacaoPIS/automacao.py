import pandas as pd
import pathlib  # verificação de extensão
### imports para a automação ###
from selenium import webdriver
from selenium.webdriver.common.by import By  # Achar os elementos
from selenium.webdriver.common.keys import Keys  # Digitação
### temporizador para carregar a pagina ###
import time
### Import para interface grafica ###
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

### Transforma em .exe ###

### Inicia o navegador ###

chrome = webdriver.Chrome("C:/Users/xavie/git/AutomacaoPIS/Automacao-PIS/chromedriver.exe")
url_abrir = 'https://forms.gle/biofPp1SwfdThVrc9'
chrome.get(url_abrir)


### Try para caso de erro ###
def error():
    messagebox.showerror(title="Extensão não suportada", message="Favor escolha um arquivo .xlsx")

### parte da seleção do arquivo ###

def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename()

### parte da Automacação###

def iniciar_aut():

    df = pd.read_excel(arquivo)
    df.dropna(subset=['PIS'], inplace=True)

    for index, column in df.iterrows():
        time.sleep(5)

        elemento_PIS = chrome.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/input')
        elemento_botao = chrome.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span')  # Elemento que seleciona o XPATH do botão

        elemento_PIS.send_keys(column["PIS"])
        elemento_botao.click()

        chrome.execute_script("window.history.go(-1)")

### interface gráfica ###

# Definição da Janela
janela_aut = Tk()
janela_aut.title("Automação para inserção de dados na web")
janela_aut.geometry("663x220+610+153")
janela_aut.resizable(width=1, height=1)

# Import imagens
img_fundo = PhotoImage(file='C:/Users/xavie/git/AutomacaoPIS/Automacao-PIS/img_fundo2.png')
bot_fund = PhotoImage(file='C:/Users/xavie/git/AutomacaoPIS/Automacao-PIS/bot_fundo.png')

# Labels
lab_fundo = Label(janela_aut, image=img_fundo)
lab_fundo.pack()

# Caixa de Erro e Aviso

# criação do botão
bt_iniciar = Button(janela_aut, image=bot_fund, command=iniciar_aut)
bt_iniciar.place(width=68, height=40, x=295, y=180)

bt_selecionar = Button(janela_aut, image=bot_fund, command=selecionar_arquivo)
bt_selecionar.place(width=68, height=40, x=195, y=180)
janela_aut.mainloop()

#teste bobo#