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

ie = webdriver.Ie("IEDriverServer.exe")
url_abrir = 'https://conectividadesocialv2.caixa.gov.br/sicns/?token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOnsidHBJbnNjcmljYW8iOiIxIiwiY29kaWdvU2VyaWFsIjoiMDAwMDAwMDAyNDYxNjg2RTlENjVCOEYzRDlGRTI0QjREQzFDQzE3OCIsImNvZFNlcmlhbEhleGFJbnZlcnQiOiI3OEMxMUNEQ0I0MjRGRUQ5RjNCODY1OUQ2RTY4NjEyNCIsInJhemFvU29jaWFsIjoiVkFMRSBWRVJERSBFTVBSRUVORElNRU5UT1MgQUdSSUNPTEFTIExUREEgRU0gUkVDVSIsImluc2NyaWNhbyI6IjAyNDE0ODU4MDAwMTI4IiwicmVzcG9uc2F2ZWwiOiJFRFVBUkRPIEpPU0UgREUgRkFSSUFTIiwiY3BmUmVzcG9uc2F2ZWwiOiIxNzQ2OTQyMjQwNCIsImNvZGlnb0NlcnRpZmljYWRvIjpudWxsLCJjZXJ0IjpudWxsfSwiY3JlYXRlZCI6MTY1MjcwMDAxMjU3MiwiZXhwIjoxNjUyNzA0ODEyfQ.bvf9Zvcdwoxp9y74xBTeDbFJpBQLmnkPvEZo0plx-wo'
ie.get(url_abrir)


### Try para caso de erro ###
def error():
    messagebox.showerror(title="Extensão não suportada", message="Favor escolha um arquivo .xlsx")

### parte da seleção do arquivo ###

def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename() # o Arquivo deve conter no minimo mais de uma coluna, sendo que a coluna que será pesquisada deve se chamar "PIS"

### parte da Automação ###

def iniciar_aut():
    
    df = pd.read_excel(arquivo)
    df.dropna(subset=['PIS'], inplace=True)

    for index, row in df.iterrows():
        time.sleep(5)

        elemento_PIS = ie.find_element(By.XPATH, '//*[@id="txtPIS"]')
        elemento_botao = ie.find_element(By.XPATH, '/html/body/form/table[2]/tbody/tr[2]/td[3]/table[3]/tbody/tr[9]/td/a[1]/img')  # Elemento que seleciona o XPATH do botão

        elemento_PIS.clear
        time.sleep(1)
        elemento_PIS.send_keys(row["PIS"])
        elemento_botao.click()

        time.sleep(3)
        ie.execute_script("window.history.go(-1)")

### interface gráfica ###

# Definição da Janela
janela_aut = Tk()
janela_aut.title("Automação para inserção de dados na web")
janela_aut.geometry("663x220+610+153")
janela_aut.resizable(width=1, height=1)

# Import imagens
img_fundo = PhotoImage(file='Automacao-PIS/img_fundo2.png')
bot_fund = PhotoImage(file='Automacao-PIS/bot_fundo.png')
bot_iniciar = PhotoImage(file='Automacao-PIS/bot_iniciar.png')

# Labels
lab_fundo = Label(janela_aut, image=img_fundo)
lab_fundo.pack()

# Caixa de Erro e Aviso

# criação do botão
bt_iniciar = Button(janela_aut, image=bot_iniciar, command=iniciar_aut)
bt_iniciar.place(width=68, height=40, x=295, y=180)

bt_selecionar = Button(janela_aut, image=bot_fund, command=selecionar_arquivo)
bt_selecionar.place(width=68, height=40, x=195, y=180)
janela_aut.mainloop()
