#!pip install selenium
#!pip install pandas as pd
#!pip install pynput

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium import webdriver
from pynput.keyboard import Key, Controller
import pandas as pd
import urllib
import time

contato_df = pd.read_excel("Lista.xlsx")
#abre o whatsapp
navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com")
midia = "teste.png"

wait = WebDriverWait(navegador, 60)
elemento_side = wait.until(EC.presence_of_element_located((By.ID, "side")))

contador = 0
Mensagens_enviadas = 0
Mensagens_falhadas = 0
nome = ['']
numero_lista = [1]

#faz um loop para cada contato da planilha
for i, mensagem in enumerate(contato_df['Mensagem']):
    try:
        pessoa = contato_df.loc[i, "Pessoa"]
        numero = contato_df.loc[i, "Numero"]
        texto = urllib.parse.quote(f"Olá {pessoa}! Na paz? \n {mensagem}")
        #abre a conversa
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        navegador.get(link)
        wait = WebDriverWait(navegador, 60)
        #if Mensagens_enviadas == 0:
        #    time.sleep(30)
        #else:
        #    time.sleep(20)
        #aguarda terminar de carregar a conversa // risco de timeout
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(text(), 'Iniciando conversa')]")))
        time.sleep(10)
        navegador.find_element("xpath", '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/div').click()
        time.sleep(5)
        attach = navegador.find_element("xpath", '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/span/div/ul/div/div[2]/li')#.send_keys(midia)
        attach.click()
        time.sleep(3)
        keyboard = Controller()
        keyboard.type("C:\\Users\\lucas\\Desktop\\1\\Bot_Mensageria\\teste.png")
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
       # time.sleep(1)
        #attach.send_keys("teste.png")
        time.sleep(5)
        send = navegador.find_element("xpath", '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
        send.click()
        #envia a mensagem
        elemento_mensagem = navegador.elemento = navegador.find_element("xpath", '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div[2]/div/p[2]')
        elemento_mensagem.send_keys(Keys.ENTER)
        #aguarda 2seg
        time.sleep(2)
        contador = contador + 1
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao enviar mensagem para o contato {pessoa} - {numero}")
        nome.append(pessoa)
        numero_lista.append(numero)
        Mensagens_falhadas = Mensagens_falhadas + 1
        continue

try:
  dados = { 'Pessoa': nome,
  'Numero': numero_lista}

  df = pd.DataFrame(dados)

  Nome_Arquivo = 'Falha_envio.xlsx'

    # Salvando o DataFrame como um arquivo XLSX
  df.to_excel(Nome_Arquivo, index=False)
  af = pd.read_excel(Nome_Arquivo)
except FileNotFoundError:
    print("Falha ao tentar criar a tabela de Falhas")

print("Finalizado")
print(f"Foram enviadas {Mensagens_enviadas} Mensagens")
print(f"Houve falha no envio de {Mensagens_falhadas} Mensagens")
navegador.quit()