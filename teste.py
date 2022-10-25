from playwright.sync_api import Page, expect
from playwright.sync_api import sync_playwright
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

### xpath a serem usados ###
xpath_input = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/input'
xpath_button = '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span'

### Seleção de arquivo ###

def selecionar_arquivo():
    global arquivo 
    arquivo = filedialog.askopenfilename()

### Automação ###

def iniciar_aut():
    global caixa_entrada
    global botao
    df = pd.read_excel(arquivo)
    df.dropna(subset=['PIS'], inplace=True)

    for index, row in df.iterrows():
        with sync_playwright() as p:
            preencher = row["PIS"]
            navegador = p.chromium.launch(headless=False) 
            pagina = navegador.new_page()
            pagina.goto("https://forms.gle/biofPp1SwfdThVrc9")
            caixa_entrada = pagina.locator(f"xpath={xpath_input}")
            expect(caixa_entrada).to_be_visible
            caixa_entrada.fill(preencher)
            botao = pagina.locator(f"xpath={xpath_button}")


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