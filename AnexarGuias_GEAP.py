import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from abc import ABC
import pandas as pd
import time
import pyautogui
from tkinter import filedialog

class PageElement(ABC):
    def __init__(self,webdriver,url=''):
        self.webdriver = webdriver
        self.url = url
    def open(self):
        self.webdriver.get(self.url)

class Login(PageElement):
    multiusuario = (By.XPATH,"/html/body/div[3]/div[3]/div/form/div[1]/label")
    login = (By.XPATH,"/html/body/div[3]/div[3]/div/form/div[2]/div[1]/div/input")
    senha = (By.XPATH,"/html/body/div[3]/div[3]/div/form/div[3]/input")
    cpf = (By.XPATH,"/html/body/div[3]/div[3]/div/form/div[2]/div[2]/div/input")
    logar = (By.XPATH,"/html/body/div[3]/div[3]/div/form/div[4]")

    def exe_login(self, login, senha, cpf):
        self.webdriver.find_element(*self.multiusuario).click()
        self.webdriver.find_element(*self.login).send_keys(login)
        self.webdriver.find_element(*self.senha).send_keys(senha)
        self.webdriver.find_element(*self.cpf).send_keys(cpf)
        self.webdriver.find_element(*self.logar).click()
     
class caminho(PageElement):
    guia = (By.XPATH,'//*[@id="main"]/div/div/div/table/tbody/tr[2]/td[7]/a')
    alerta = (By.XPATH,' /html/body/div[2]/div/center/a')


    
    def exe_caminho(self):
        time.sleep(4)
        try:
            self.webdriver.find_element(*self.alerta).click()
        except:
            print('Alerta não apareceu')
        webdriver.get("https://www2.geap.com.br/PRESTADOR/tiss-detalhamento-de-protocolo.asp?NroProtocolo=9651213#")
        self.webdriver.find_element(*self.guia).click()

class Anexar_Guia(PageElement):
    anexar = (By.XPATH,'//*[@id="flUpload"]')
    salvar = (By.XPATH,'//*[@id="MenuOptionUpdate"]')
    confere = (By.XPATH,'//*[@id="grvDocumentos_ctl02_ctl00"]')

    def injetar_guia(self):
        count = 0
        faturas_df = pd.read_excel(planilha)
        for index, linha in faturas_df.iterrows():
            if (f"{linha['Guia Anexada']}") == "Sim":
                print(f"{linha['Guia Anexada']}")
                count = count + 1
                continue
            else:
                print('Guia pronta para ser anexada')
            count = count + 1
            print('------------------------------------------------------------------------------------------------------')
            print(f"{linha['Nro Guia GEAP']}")
            webdriver.get("https://www2.geap.com.br/digitaTiss/UploadDocumentosXML.aspx?numeroGuiaOperadora=" + f"{linha['Nro Guia GEAP']}")
            print('Entrando na guia')
            time.sleep(2)
            self.webdriver.find_element(*self.anexar).send_keys(linha["Caminho"])
            print('Guia Anexada')
            time.sleep(2)
            element = WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="MenuOptionUpdate"]')))
            self.webdriver.find_element(*self.salvar).click()
            time.sleep(2)
            print('Salvo')
            print(count, 'Guia(s) Anexada(s)')

            
            element = WebDriverWait(webdriver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="grvDocumentos_ctl02_ctl02"]')))
            print('Achou o elemento')
            anexado = webdriver.find_element(By.XPATH,'//*[@id="grvDocumentos_ctl02_ctl02"]').text
            print(anexado)
            if anexado == 'application/pdf':
                dados = ['Sim']
                df = pd.DataFrame(dados)
                print('Conferencia feita')
                book = load_workbook(planilha)
                writer = pd.ExcelWriter(planilha, engine='openpyxl')
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                df.to_excel(writer, 'Planilha1', startrow= count, startcol=3, header=False, index=False)

                writer.save()
                continue
            
            
     
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
login_usuario = "lucas.timoteo"
senha_usuario = "Caina9018"


planilha = filedialog.askopenfilename()

url = "https://www2.geap.com.br/auth/prestador.asp"

webdriver = webdriver.Chrome(r"\\10.0.0.71\atualiza\teste-python\chromedriver.exe")

login_page = Login(webdriver, url)
login_page.open()

webdriver.maximize_window()
time.sleep(4)
pyautogui.write(login_usuario)
pyautogui.press("TAB")
time.sleep(1)
pyautogui.write(senha_usuario)
pyautogui.press("enter")
time.sleep(4)

login_page.exe_login(
    login = "23003723",
    senha = "amhpdf0073",
    cpf = "66661692120"
 )

time.sleep(4)

caminho(webdriver, url).exe_caminho()

webdriver.switch_to.window(webdriver.window_handles[0])

time.sleep(2)

Anexar_Guia(webdriver, url).injetar_guia()

