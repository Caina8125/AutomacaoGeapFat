import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
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
    alerta = (By.XPATH,' /html/body/div[2]/div/center/a')
    guia = (By.XPATH,'//*[@id="objTableDetalhe"]/tbody/tr[3]/td[1]/a')
    envio_xml = (By.XPATH,'//*[@id="main"]/div/div/div[2]/div[2]/article/div[4]/div[4]/div[4]/div[4]/div/div[2]/ul/li[2]/a')
    sem_erros = (By.XPATH,'//*[@id="StaErro"]/option[2]')
    listar = (By.XPATH,'//*[@id="MenuOptionReport"]')

    def exe_caminho(self):
        time.sleep(4)
        try:
            self.webdriver.find_element(*self.alerta).click()
        except:
            print('Alerta não apareceu')

        webdriver.get("https://www2.geap.com.br/PRESTADOR/portal-tiss.asp#")
        time.sleep(2)
        self.webdriver.find_element(*self.envio_xml).click()
        time.sleep(1)
        webdriver.switch_to.window(webdriver.window_handles[1])
        time.sleep(1)
        self.webdriver.find_element(*self.sem_erros).click()
        time.sleep(1)
        self.webdriver.find_element(*self.listar).click()


class Anexar_Guia(PageElement):
    anexar = (By.XPATH,'//*[@id="fupDoc"]')
    adicionar = (By.XPATH,'//*[@id="btnAdicionar"]')

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
            webdriver.get("https://www2.geap.com.br/PRESTADOR/auditoriadigital/rpt/DetalhamentoGuia.aspx?IdGsp=" + f"{linha['Nro Guia GEAP']}")
            print('Entrando na guia')
            time.sleep(2)
            self.webdriver.find_element(*self.anexar).send_keys(linha["Caminho"])
            time.sleep(1)
            self.webdriver.find_element(*self.adicionar).click()
            time.sleep(2)
            print('Guia Anexada')

#------------------------------------------------------------------------------------------------------------------------------------------------

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

time.sleep(2)

Anexar_Guia(webdriver, url).injetar_guia()