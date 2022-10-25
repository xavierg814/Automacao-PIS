from email import message
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
from selenium.common.exceptions import *

### Transforma em .exe ###

### Inicia o navegador ###
try:
    chrome = webdriver.Chrome("chromedriver.exe")
    url_abrir = 'https://forms.gle/biofPp1SwfdThVrc9'
    chrome.get(url_abrir)
except SessionNotCreatedException as erro:
    print("Versão do Drivre está incorreta, favor atualizar driver dentro da pasta da automação")
    messagebox.showerror("Erro", "Versão do chromedrive é incompativél com a versão atual do chrome, favor verificar sua ver~sao do chrome e atualize o chromedriver")

### parte da seleção do arquivo ###

def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename() # o Arquivo deve conter no minimo mais de uma coluna, sendo que a coluna que será pesquisada deve se chamar "PIS"

### parte da Automacação###

def iniciar_aut():
    try:
        df = pd.read_excel(arquivo)
        df.dropna(subset=['PIS'], inplace=True)

        for index, row in df.iterrows():
            time.sleep(2)

            elemento_PIS = chrome.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/input')
            elemento_botao = chrome.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span')  # Elemento que seleciona o XPATH do botão

            elemento_PIS.clear()
            time.sleep(1)
            elemento_PIS.send_keys(row["PIS"])
            elemento_botao.click()

            time.sleep(3)
            chrome.execute_script("window.history.go(-1)")
        messagebox.showinfo("Sucesso", "Todos os PIS da coluna foram enviados")

    except NoSuchElementException as erro:
        print("Xpath não encontrado, verifique se está na tela correta ou relate o problema ao dev", erro)
        messagebox.showerror("Erro", "Falha ao Encontrar o Xpath", erro)
    except KeyError as erro:
        print("Coluna PIS não encontrada, favor verificar nome da coluna no arquivo excel", erro)
        messagebox.showerror("Erro", "Coluna PIS não encontrada, favor verificar planilha selecionada")
    except NoSuchWindowException as erro:
        print("Janela alvo da automação foi fechada, favor reinicie o programa", erro)
        messagebox.showerror("Erro", "Janela Alvo foi fehcada, favor reinicie a automação")
    except Exception as erro:
        print("Um erro Inesperado ocorreu, favor entrar em contato com o dev", erro)
        messagebox.showerror("Erro", f"Um Erro inesperado ocorreu, favor entrar em contato com o dev: {erro}")

### interface gráfica ###

# Definição da Janela
janela_aut = Tk()
janela_aut.title("Automação para inserção de dados na web")
janela_aut.geometry("663x220+610+153")
janela_aut.resizable(width=1, height=1)

# Import imagens
img_fundo = PhotoImage(file='img_fundo2.png')
bot_fund = PhotoImage(file='bot_fundo.png')
bot_iniciar = PhotoImage(file='bot_iniciar.png')

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