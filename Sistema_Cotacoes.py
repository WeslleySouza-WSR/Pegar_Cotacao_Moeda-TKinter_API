import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkcalendar import DateEntry
from datetime import datetime
import requests
import pandas as pd
import numpy as np


def procurar():
    global caminho_arq
    caminho_arq = askopenfilename(title='Selecione um arquivo excel')
    if caminho_arq:
        caminho_selecionado['text'] = f'Arquivo Selecionado: {caminho_arq}'
    else:
        caminho_selecionado['text'] = f'Nenhum arquivo selecionado'


def atualizar():
    try:
        if caminho_arq:
            df = pd.read_excel(caminho_arq)
            df_moedas = df.iloc[:, 0]
            dia_ini, mes_ini, ano_ini = calendario_moeda_inicial.get().split('/')
            dia_fin, mes_fin, ano_fin = calendario_moeda_final.get().split('/')
            dias = (calendario_moeda_final.get_date() - calendario_moeda_inicial.get_date()).days
            for moeda in df_moedas:
                requisicao = requests.get(f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/{dias + 1}'
                                          f'?start_date={ano_ini}{mes_ini}{dia_ini}&end_date='
                                          f'{ano_fin}{mes_fin}{dia_fin}').json()
                for dia in requisicao:
                    timestamp = int(dia['timestamp'])
                    bid = dia['bid']
                    data = datetime.fromtimestamp(timestamp)
                    df_dia, df_mes, df_ano = data.day, data.month, data.year
                    if f'{df_dia}/{df_mes}/{df_ano}' not in df:
                        df[f'{df_dia}/{df_mes}/{df_ano}'] = np.nan
                    df.loc[df.iloc[:, 0] == moeda, f'{df_dia}/{df_mes}/{df_ano}'] = bid
            df.to_excel('teste.xlsx')
            mensagem_atualizado['text'] = 'Arquivo atualizado com sucesso.'
    except:
        mensagem_atualizado['text'] = 'Selecione um arquivo em Excel no formato correto'


def pegar_cotacao():
    dia, mes, ano = calendario_moeda.get().split('/')
    requisicao = requests.get(f'https://economia.awesomeapi.com.br/json/daily/{moeda.get()}-'
                              f'BRL?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}').json()
    if moeda.get() and calendario_moeda.get():
        mensagem_cotacao_moeda['text'] = f'A Cotação da moeda {moeda.get()} no dia {calendario_moeda.get()} ' \
                                         f'R$ {float(requisicao[0]["bid"]):.2f} '


dic_moedas = requests.get('https://economia.awesomeapi.com.br/json/all').json()
moedas = list(dic_moedas.keys())

janela = tk.Tk()
janela.title('Ferramenta de Cotação de Moedas')
janela.rowconfigure(list(range(11)), weight=1)
janela.columnconfigure(list(range(3)), weight=1)

tk.Label(text='Cotação de 1 Moeda Específica', borderwidth=2, relief='solid').grid(row=0, column=0, columnspan=3,
                                                                                   sticky='NSEW', pady=1, padx=0.5)
tk.Label(text='Selecione a Moeda Desejada:', anchor='e').grid(row=1, column=0, columnspan=2, sticky='NSEW', pady=1,
                                                              padx=0.5)
moeda = ttk.Combobox(janela, values=moedas)
moeda.grid(row=1, column=2, sticky='NSEW', pady=1, padx=0.5)
tk.Label(text='Selecione o Dia da Cotação:', anchor='e').grid(row=2, column=0, columnspan=2, sticky='NSEW', pady=1,
                                                              padx=0.5)
calendario_moeda = DateEntry(year=2022, locale='pt_br')
calendario_moeda.grid(row=2, column=2, sticky='NSEW', pady=1, padx=0.5)
mensagem_cotacao_moeda = tk.Label(text='')
mensagem_cotacao_moeda.grid(row=3, column=0, columnspan=2, sticky='NSEW', pady=1, padx=0.5)
tk.Button(text='Pegar Cotação', command=pegar_cotacao).grid(row=3, column=2, sticky='NSEW', pady=1, padx=0.5)

tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid').grid(row=4, column=0, columnspan=3,
                                                                                 sticky='NSEW', pady=1, padx=0.5)
tk.Label(text='Selecione um Arquivo em Excel com as Moedas na Coluna A:').grid(row=5, column=0, columnspan=2,
                                                                               sticky='NSEW', pady=1, padx=0.5)
tk.Button(text='Selecionar arquivo', command=procurar).grid(row=5, column=2, sticky='NSEW', pady=1, padx=0.5)
caminho_selecionado = tk.Label(text='Nenhum arquivo selecionado', anchor='e')
caminho_selecionado.grid(row=6, column=0, columnspan=3, sticky='NSEW', pady=1, padx=0.5)
tk.Label(text='Data Inicial:', anchor='e').grid(row=7, column=0, columnspan=1, sticky='NSEW', pady=1, padx=0.5)
calendario_moeda_inicial = DateEntry(year=2022, locale='pt_br')
calendario_moeda_inicial.grid(row=7, column=1, columnspan=1, sticky='NSEW', pady=1, padx=0.5)
tk.Label(text='Data Final:', anchor='e').grid(row=8, column=0, columnspan=1, sticky='NSEW', pady=1, padx=0.5)
calendario_moeda_final = DateEntry(year=2022, locale='pt_br')
calendario_moeda_final.grid(row=8, column=1, columnspan=1, sticky='NSEW', pady=1, padx=0.5)
tk.Button(text='Atualizar Cotações', command=atualizar).grid(row=9, column=0, sticky='NSEW', pady=8, padx=0.5)
mensagem_atualizado = tk.Label(text='')
mensagem_atualizado.grid(row=9, column=1, columnspan=3, sticky='NSEW', pady=8, padx=0.5)

janela.mainloop()
