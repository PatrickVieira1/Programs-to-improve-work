import re
import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pymsgbox as pymsgbox
import time
import pyautogui

iframe = "//body/div[2]/div[2]/div[1]/div[4]/iframe[1]"
GerarEstrutura = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div/div/div/div[2]/div/div[2]/div/div[2]/div/div/div[6]/div/div[2]/div/div/div/div/div[2]/div/label'
Item = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[2]/div[1]/div/div[2]/div/div/div[6]/input'
OV = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[2]/div[1]/div/div[2]/div/div/div[9]/input'
Gerar = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[2]/div[1]/div/div[1]/div[2]/div[1]/div[1]/label'
CampoBserp = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[1]/div[4]/div/div/input'
CampoHserp = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[2]/div[4]/div/div/input'
CampoCabeceira = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[3]/div[4]/div/div/input'
CampoRows = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[5]/div[4]/div/div/input'
CampoCircuitos = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[6]/div[4]/div/div/input'
CampoAletas = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[8]/div[4]/div/div/input'
CampoDiametroTubos = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[9]/div[4]/div/div/input'
CampoConexao = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[10]/div[4]/div/div/input'
CampoHidraulica = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[2]/div/div[4]/div/div[2]/div[7]/div/div/div[3]/div/div[11]/div[4]/div/div/input'
SalvareSair = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div[1]/div/div[1]/div[2]/div[1]/div[1]'
LabelPrecoProduto = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[4]/div/div/div[3]/div[2]'
LabelGerarEstrutura = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[3]/div/div/div[3]/div[1]'

CodigoVariante = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[2]/div[1]/div/div[2]/div/div/div[3]/input'
CodigoDescricao = '//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[2]/div[1]/div/div[2]/div/div/label[2]'

RecuperarVariante ='//body/div[3]/div[2]/div/div[3]/div/div[4]/div/div[2]/div[1]/div/div[2]/div/div/div[1]/div/svg'


username = pymsgbox.prompt('Qual seu usuário do LN?', default='')
password = pymsgbox.password('Qual sua senha do LN?', default='')


lnURL = 'ln.troxbrasil.com.br:8312/webui/servlet/standalone'

CodigoOV = 'BR1333459'
  
def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Selecione um arquivo",
                                          filetypes = (("Text files",
                                                        "*.txt*"),
                                                       ("all files",
                                                        "*.*")))
      
    # Change label contents
    label_file_explorer.configure(text="Arquivo Aberto: "+os.path.basename(filename))

window = tk.Tk()
window.title("Comprador automático de serpentina")

frame1 = tk.Frame(borderwidth=3,relief = "raised")
frame1.grid(column=0,row=0)

CabeceiraFrame = tk.LabelFrame(text="Cabeceira",borderwidth=3)
CabeceiraFrame.place(x=10,y=0)



var1 = tk.StringVar()
var2 = tk.StringVar()
var1.set(0)
var2.set(0)


r1 = tk.Radiobutton(CabeceiraFrame, text=("Aço Inox"), value = 1, variable=var1)
r2 = tk.Radiobutton(CabeceiraFrame, text=("Alumínio"), value = 2, variable=var1)
r1.grid()
r2.grid()

HidraulicaFrame = tk.LabelFrame(borderwidth = 3, text="Lado Hidraulica")
HidraulicaFrame.place(x=100,y=0)

r3 = tk.Radiobutton(HidraulicaFrame, text="Direita", value = 1, variable=var2)
r4 = tk.Radiobutton(HidraulicaFrame, text="Esquerda", value = 2, variable=var2)
r3.grid()
r4.grid()

FilesFrame = tk.LabelFrame(text="Arquivos a selecionar",borderwidth=3)
FilesFrame.place(x=200,y=0)


label_file_explorer = Label(FilesFrame,
                            text = "Selecione o arquivo serpentina",
                            width = 50, height = 4,
                            fg = "blue", justify='left')
  
      
button_explore = Button(FilesFrame,
                        text = "Pesquisar Arquivo",
                        command = browseFiles)

label_file_explorer.grid(column = 0, row = 1)
button_explore.grid(column = 0, row = 2)

ExecutarFrame = tk.LabelFrame(borderwidth=0)
ExecutarFrame.place(x=0,y=85)
count = 0
Show = 0
ListaCodigosArquivos = []
ListaCodigosVariantes = []
ListaCodigosDescricao = []

def BotaoCounter():

    global count
    count += 1

def DadosSerp():
    arquivo_serp = filename
    with open(arquivo_serp) as f:
        for line in f:

            Comprimento = re.search(r'(?<=Comprimento:).*(.+?),',line)
            if Comprimento:
                Bserp = Comprimento.group(0).replace(" ", "").replace(",","")

            Altura = re.search(r'(?<=Altura:).*(.+?),',line)
            if Altura:
                Hserp = Altura.group(0).replace(" ", "").replace(",","")
                if Hserp == '1066':
                    Hserp = '1067'
                elif Hserp == '609':
                    Hserp = '610'
                elif Hserp == '914':
                    Hserp = '915'

            Rows = re.search(r'''(?<=(Rows)).*(.+?)''',line)
            if Rows:
                RowsSerp = Rows.group(0).replace(" ", "").replace("(fileiras):","")

            Circuitos = re.search(r'''(?<=(Circuitos:)).*(.+?)''',line)
            if Circuitos:
                CircuitosSerp = Circuitos.group(0).replace(" ", "")

            Aletas = re.search(r'''(?<=(Aletas po UC:)).*(.+?)\d''',line)
            if Aletas:
                AletasSerp = Aletas.group(0).replace(" ", "")
                if AletasSerp == '3,93':
                    AletasSerp = '10'
                elif AletasSerp == '3,15' or AletasSerp == '3,14':
                    AletasSerp = '8'

            Coletor = re.search(r'''(?<=(Conexão:)).*(.+?)\d''',line)
            if Coletor:
                ColetorSerp = Coletor.group(0).replace(" ", "")
                if ColetorSerp == '2,00':
                    ColetorSerp = '2'
                elif ColetorSerp == '2,50':
                    ColetorSerp = '2 1/2'
                elif ColetorSerp == '1,50':
                    ColetorSerp = '1 1/2'
                elif ColetorSerp == '1,25':
                    ColetorSerp = '1 1/4'
                elif ColetorSerp == '1,00':
                    ColetorSerp = '1'
                elif ColetorSerp == '0,75':
                    ColetorSerp = '3/4'
                elif ColetorSerp == '0,50':
                    ColetorSerp = '1/2'
                elif ColetorSerp == '3,00':
                    ColetorSerp = '3'

    if int(var1.get()) == 1:
        CabeceiraValue = 'I4'
    elif int(var1.get()) ==2:
        CabeceiraValue = 'AL'

    if int(var2.get()) == 1:
        HidraulicaValue = 'D'
    elif int(var2.get()) == 2:
        HidraulicaValue = 'E'


    driver = webdriver.Firefox(executable_path="U:\\Engenharia\\Usuários\\Patrick.Vieira\\geckodriver.exe")
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    time.sleep(5)
    driver.get(lnURL)

    pyautogui.write(username)
    pyautogui.press('tab')
    pyautogui.write(password)
    pyautogui.press('enter')


    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, iframe)))
    driver.switch_to.frame(driver.find_element_by_xpath(iframe))
    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, GerarEstrutura)))
    driver.find_element_by_xpath(GerarEstrutura).click()

    #Espera sair a <div class="ListItem">
    wait.until(ec.invisibility_of_element_located((By.XPATH,
              "//div[@class='ListItem']")))
    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, Item)))
    driver.find_element_by_xpath(Item).click()
    driver.find_element_by_xpath(Item).send_keys("EF387221")
    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, OV)))
    driver.find_element_by_xpath(OV).click()
    driver.find_element_by_xpath(OV).send_keys(CodigoOV)
    driver.find_element_by_xpath(Gerar).click()

    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, CampoBserp)))
    driver.find_element_by_xpath(CampoBserp).click()
    driver.find_element_by_xpath(CampoBserp).send_keys(Bserp)
    driver.find_element_by_xpath(CampoBserp).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoHserp)))
    #driver.find_element_by_xpath(CampoHserp).click()
    driver.find_element_by_xpath(CampoHserp).send_keys(Hserp)
    driver.find_element_by_xpath(CampoHserp).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoCabeceira)))
    driver.find_element_by_xpath(CampoCabeceira).click()
    driver.find_element_by_xpath(CampoCabeceira).send_keys(CabeceiraValue)
    driver.find_element_by_xpath(CampoCabeceira).send_keys(Keys.ENTER)

    pyautogui.press('tab')

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoRows)))
    driver.find_element_by_xpath(CampoRows).click()
    driver.find_element_by_xpath(CampoRows).send_keys(RowsSerp)
    driver.find_element_by_xpath(CampoRows).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoCircuitos)))
    driver.find_element_by_xpath(CampoCircuitos).click()
    driver.find_element_by_xpath(CampoCircuitos).send_keys(CircuitosSerp)
    driver.find_element_by_xpath(CampoCircuitos).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoAletas)))
    driver.find_element_by_xpath(CampoAletas).click()
    driver.find_element_by_xpath(CampoAletas).send_keys(AletasSerp)
    driver.find_element_by_xpath(CampoAletas).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoDiametroTubos)))
    driver.find_element_by_xpath(CampoDiametroTubos).click()
    driver.find_element_by_xpath(CampoDiametroTubos).send_keys("5/8")
    driver.find_element_by_xpath(CampoDiametroTubos).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoConexao)))
    driver.find_element_by_xpath(CampoConexao).click()
    driver.find_element_by_xpath(CampoConexao).send_keys(ColetorSerp)
    driver.find_element_by_xpath(CampoConexao).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CampoHidraulica)))
    driver.find_element_by_xpath(CampoHidraulica).click()
    driver.find_element_by_xpath(CampoHidraulica).send_keys(HidraulicaValue)
    driver.find_element_by_xpath(CampoHidraulica).send_keys(Keys.ENTER)

    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, SalvareSair)))
    driver.find_element_by_xpath(SalvareSair).click()

    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, LabelPrecoProduto)))
    driver.find_element_by_xpath(LabelPrecoProduto).click()

    WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, LabelGerarEstrutura)))
    driver.find_element_by_xpath(LabelGerarEstrutura).click()

    CodigoVar = WebDriverWait(driver, 30).until(ec.element_to_be_clickable((By.XPATH, CodigoVariante))).get_attribute('value')
    CodigoDes = WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, CodigoDescricao))).text

    ListaCodigosArquivos.append(os.path.split(os.path.basename(filename))[1])
    ListaCodigosVariantes.append(CodigoVar)
    ListaCodigosDescricao.append(CodigoDes)
    SalvarBotao.grid(column=0,row=0,padx=10)
    driver.close()

def SalvarPlanilha():
    messagebox.showinfo('Por favor', 'Escolha onde quer salvar a planilha com os códigos')
    localSalvo = filedialog.askdirectory()
    localSalvoCompleto = localSalvo + '/' + 'SERPENTINAS GERADAS - '+ CodigoOV + '.xlsx'
    localSalvoPython = localSalvoCompleto.replace('/', '//')
    dicionario = {'Nome arquivo': ListaCodigosArquivos, 'Codigos': ListaCodigosVariantes, 'Descrição':ListaCodigosDescricao}
    df = pd.DataFrame(dicionario)
    writer = pd.ExcelWriter(localSalvoPython, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='DESENHOS', index=False)
    writer.save()
    messagebox.showinfo('Sucesso!', 'Arquivo Salvo com êxito!')



ExecutarBotao = tk.Button(ExecutarFrame, text ='Executar', command= DadosSerp)
ExecutarBotao.grid(column=1,row=0,padx=40)

SalvarBotao = tk.Button(ExecutarFrame, text ='Salvar Planilha', command= SalvarPlanilha)

window.geometry("600x140+300+400")
#window.eval('tk::PlaceWindow . center')
#window.configure(background='black')

tk.mainloop()