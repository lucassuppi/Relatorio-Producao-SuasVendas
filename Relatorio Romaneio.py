from datetime import datetime
import ctypes
import sys
  
LOGIN = seu email
SENHA = sua senha
  
def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)
Mbox('AVISO', 'Gerar relatório de produção?', 1)
Mbox('AVISO', 'Aguarde!', 1)
  
import time
time.sleep(2)

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
navegador = webdriver.Chrome()
navegador.get("https://www.suasvendas.com/forms/login.aspx")
navegador.set_window_position(2000,2000)
navegador.find_element_by_xpath('//*[@id="txtEmail"]').send_keys(LOGIN)
navegador.find_element_by_xpath('//*[@id="txtSenha"]').send_keys(SENHA) 
navegador.find_element_by_xpath('//*[@id="btnAcessar"]').send_keys(Keys.RETURN)
 
time.sleep(10)

navegador.find_element_by_xpath('//*[@id="menuCollapse"]/ul/li[13]/a').click()
navegador.find_element_by_xpath('//*[@id="ystab-pedidos-tab"]').click()
navegador.execute_script("window.open('https://app3.suasvendas.com/Modulo/YourSales/Relatorios/V2/RomaneioDeCargaV2.aspx', '_blank')")

navegador.implicitly_wait(5)
navegador.switch_to_window(navegador.window_handles[1])

navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_drpPeriodo"]').click()
navegador.find_element_by_xpath('//*[@id="APeriodoDigitar"]').click()

from selenium.webdriver.common.keys import Keys

navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_txtPeriodoDataInicial"]').send_keys("01/01/2020")

import time
time.sleep(5)

navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_txtPeriodoDataInicial"]').send_keys(Keys.TAB)

navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_txtPeriodoDataFinal"]').click()

import time
time.sleep(5)

from datetime import date

data_atual = date.today()
data_em_texto = data_atual.strftime('%d/%m/%Y')

navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_txtPeriodoDataFinal"]').send_keys(data_em_texto)
navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_btnPeriodoDigitarDatas"]').click()
navegador.find_element_by_xpath('//*[@id="ctl00_Periodo_BtnPeriodoOK"]').click()
navegador.find_element_by_xpath('//*[@id="ctl00_TdFiltroMaster"]/span[2]/a').click()
navegador.find_element_by_xpath('//*[@id="ctl00_CPHFiltroRelatorio_StatusPedido_CblFiltro"]/tbody/tr[3]/td/span').click()
navegador.find_element_by_xpath('//*[@id="ctl00_CPHFiltroRelatorio_StatusPedido_BtnFilter"]').click()
navegador.find_element_by_xpath('//*[@id="ctl00_CabecalhoRelatorio_LiExportarExcel"]/a').click()

import time
time.sleep(10)

def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)
Mbox('AVISO', 'Relatorio baixado, edite.', 1)

navegador.quit()