from tkinter import *
from tkinter import filedialog
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

janela = Tk()

janela.title("Bot Mensageria")
janela.geometry("450x200")
largura_janela = 400
altura_janela = 200
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

# Define a geometria da janela com a posição centralizada
janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

def selecionar_imagem():
    global file_path
    file_path = filedialog.askopenfilename()
    file_path = file_path.replace("/", "\\")
    label_localizacao_arquivo.config(text=file_path)

def fechar_janela():
    janela.destroy()

label_lerro1= Label(janela, text=(f""), font=("Arial", 5))
label_lerro1.pack(pady=5)

botao_selecionar = Button(janela, text="Selecionar Imagem", font=("Arial", 12), command=selecionar_imagem)
botao_selecionar.pack(pady=10)

label_localizacao_arquivo = Label(janela, text="Arquivo", font=("Arial", 12), )
label_localizacao_arquivo.pack(pady=5)

botao_ok = Button(janela, text="OK",  font=("Arial", 12), command=fechar_janela)
botao_ok.pack(pady=10)

janela.mainloop()

contato_df = pd.read_excel("Lista.xlsx")
#abre o whatsapp
navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com")

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
        #aguarda terminar de carregar a conversa // risco de timeout
        wait.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')))
        time.sleep(10)
        navegador.find_element("xpath", '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/div').click()
        time.sleep(3)
        attach = navegador.find_element("xpath", '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/span/div/ul/div/div[2]/li')#.send_keys(midia)
        attach.click()
        time.sleep(3)
        keyboard = Controller()
        keyboard.type(file_path)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        time.sleep(3)
        send = navegador.find_element("xpath", '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
        send.click()
        #aguarda
        time.sleep(3)
        Mensagens_enviadas = Mensagens_enviadas + 1

    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao enviar mensagem para o contato {pessoa} - {numero}")
        nome.append(pessoa)
        numero_lista.append(numero)
        dados = { 'Pessoa': nome,
        'Numero': numero_lista}
        Mensagens_falhadas = Mensagens_falhadas + 1
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

        continue

janela = Tk()
janela.title("Bot Mensageria")
janela.geometry("400x200")

largura_janela = 400
altura_janela = 200
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2

# Define a geometria da janela com a posição centralizada
janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

label_enviadas = Label(janela, text=(f"\n Foram enviadas:  {Mensagens_enviadas} Mensagens"), font=("Arial", 15), )
label_enviadas.pack(pady=5)

label_lerro= Label(janela, text=(f"Falharam:  {Mensagens_falhadas} Mensagens"), font=("Arial", 15), )
label_lerro.pack(pady=5)

botao_ok = Button(janela, text="OK", font=("Arial", 15),  command=fechar_janela)
botao_ok.pack(pady=10)

janela.mainloop()
navegador.quit()