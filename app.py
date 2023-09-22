from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
#como tem seleção dropdown precisa importar Select
from selenium.webdriver.support.select import Select 
from time import sleep

oab = 133864

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)

# entrar no site da - https://pje-consulta-publica.tjmg.jus.br

driver.get("https://pje-consulta-publica.tjmg.jus.br")

sleep(20)
# digitar númeor oab e selecionar estado
campo_oab = driver.find_element(By.XPATH,"//*[@id='fPP:Decoration:numeroOAB']")
campo_oab.send_keys(oab)
dropdown_estados = driver.find_element(By.XPATH,'//*[@id="fPP:Decoration:estadoComboOAB"]')
opcoes_estados = Select(dropdown_estados)
opcoes_estados.select_by_visible_text('SP')

# clicar em pesquisar
pesquisar = driver.find_element(By.XPATH,'//*[@id="fPP:searchProcessos"]')
pesquisar.click()
sleep(10)

# entrar em cada um dos processos

# extrair o número do processo e data de distribuição
# extrair e guardar todas a últimas movimentações 
# guardar tudo no excel, separados por processo

