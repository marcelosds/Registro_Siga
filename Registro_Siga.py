import pyautogui as pyautogui
import time
import sys
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select


tabela = pd.read_excel("Atividades_Siga.xlsx")

tabela1 = pd.read_excel("Atividades_Siga.xlsx", sheet_name="acesso")

for i, usuario in enumerate(tabela1["usuario"]):
    senha = tabela1.loc[i, "senha"]


nav = webdriver.Chrome()
nav.get("http://siga.govbr.com.br/login.php")
nav.maximize_window()


time.sleep(2)
pyautogui.typewrite(str(usuario))
pyautogui.press('tab')
pyautogui.typewrite(str(senha))
pyautogui.press('enter')
time.sleep(30)


# buscando informações ordenadas da tabela / planilha
for i, chamado in enumerate(tabela["Chamado"]):
    data1 = tabela.loc[i, "Data_Inicio"]
    datainicio = data1.strftime('%d%m%Y')
    hora1 = tabela.loc[i, "Hora_Inicio"]
    horainicio = hora1.strftime('%H:%M')
    data2 = tabela.loc[i, "Data_Fim"]
    datafim = data2.strftime("%d%m%Y")
    hora2 = tabela.loc[i, "Hora_Fim"]
    acompanhamento = tabela.loc[i, "Tipo_Acompanhamento"]

    horafim = hora2.strftime('%H:%M')
    atividade = tabela.loc[i, "Atividade"]
    # alimentando os campos

    nav.find_element('xpath', '//*[@id="dsconteudolocalizacao"]').clear()
    time.sleep(1)
    nav.find_element('xpath', '//*[@id="dsconteudolocalizacao"]').send_keys(str(chamado))
    time.sleep(1)
    nav.find_element('xpath', '//*[@id="dsconteudolocalizacao"]').send_keys(Keys.ENTER)
    time.sleep(5)
    iframe = nav.find_element('xpath', "/html/body/div/table/tbody/tr[2]/td/div/iframe")
    nav.switch_to.frame(iframe)
    WebDriverWait(nav, 20)
    nav.find_element(By.XPATH, '//*[@id="gridChamadosRow_0"]/td[3]').click()
    time.sleep(5)


    # Registro da horas e atividades
    nav.find_element(By.XPATH, '//*[@id="dtinicioacompanhamentoedit"]').click()
    nav.find_element(By.XPATH, '//*[@id="dtinicioacompanhamentoedit"]').send_keys(str(datainicio))
    nav.find_element(By.XPATH, '//*[@id="dtinicioacompanhamentoedit"]').send_keys(keys.Keys.TAB)
    nav.find_element(By.XPATH, '//*[@id="dtinicioacompanhamentoedit"]').send_keys(keys.Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH, '//*[@id="hrinicioacompanhamentoedit"]').send_keys(str(horainicio))
    nav.find_element(By.XPATH, '//*[@id="hrinicioacompanhamentoedit"]').send_keys(keys.Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH, '//*[@id="dtterminoacompanhamentoedit"]').send_keys(str(datafim))
    nav.find_element(By.XPATH, '//*[@id="dtterminoacompanhamentoedit"]').send_keys(keys.Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH, '//*[@id="hrterminoacompanhamentoedit"]').send_keys(str(horafim))
    nav.find_element(By.XPATH, '//*[@id="hrterminoacompanhamentoedit"]').send_keys(keys.Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH, '//*[@id="dsacompanhamento"]').send_keys(str(atividade))
    nav.find_element(By.XPATH, '//*[@id="dsacompanhamento"]').send_keys(keys.Keys.TAB)
    time.sleep(1)

    # Buscando itens da lista suspensa dos tipos de acompanhamento
    select_element = nav.find_element(By.XPATH, '//*[@id="cdtipoacompanhamento"]')
    select_object = Select(select_element)

    # Setando item da lista de mesmo nome que o informado na planilha
    try:
        select_object.select_by_visible_text(str(acompanhamento))
    except NoSuchElementException:
        print("Tipo de acompanhamento não existe para este tipo de chamado!")
    time.sleep(1)

    # Gravar registro
    nav.find_element(By.XPATH, '//*[@id="btnEditAcompanhamento"]').click()
    time.sleep(1)
    nav.switch_to.default_content()
    WebDriverWait(nav, 20)
    time.sleep(1)
sys.exit()
