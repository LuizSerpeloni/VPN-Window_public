# Versão 1.3

import openpyxl
import pandas as pd
from datetime import datetime
import configparser
from tkinter import *

# Processos do Parser:
# Cria o objeto ConfigParser
config = configparser.ConfigParser()

# Leia o arquivo de configuração
config.read('conf.txt', encoding='utf-8')

# Pega o caminho da planilha:
path = config['Excel_local']['local']
#####

#####
# Faz a comparação
def wizard ():
    # Abre o excel necessário
    plan_data = pd.read_excel(path)

    # Coleta as variáveis para verificar a data
    now = datetime.now()

    # Verifica se há datas vencidas e coloca em no dataframe vencidos:
    columns = ['Colaborador', 'Data Expiração']
    vencidos = plan_data[columns].where(plan_data['Data Expiração'] <= now)
    vencidos = vencidos.dropna()

    if vencidos.empty:
        texto_resposta['text'] = f'''Tudo em dia!'''
    else:
        texto_resposta['text'] = f'''Aviso: Há certificados vencidos: 
                                    {vencidos.to_string(index = False)}'''

### Janela
janela = Tk()
janela.geometry('300x400')
janela.title('Planilha VPN')
texto = Label(janela, text="Clique no botão para verificar as datas")
texto.grid(column=0, row=0, padx=10, pady=10)
# Botão
botao = Button(janela, text='Consultar', command=wizard)
botao.grid(column=0, row=1, padx=10, pady=10)
# Resposta
texto_resposta = Label(janela, text="")
texto_resposta.grid(column=0, row=2, padx=10, pady=10)
janela.mainloop()
