import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from tkinter import Tk, filedialog
import PySimpleGUI as sg


#Códigos das Ações da Bovespa que serão buscados
acoes = ['RRRP3.SA', 'ALSO3.SA', 'ALPA4.SA', 'ABEV3.SA', 'ARZZ3.SA', 'ASAI3.SA', 'AZUL4.SA',
          'B3SA3.SA', 'BPAN4.SA', 'BBSE3.SA', 'BBDC3.SA', 'BBDC4.SA', 'BRAP4.SA', 'BBAS3.SA',
            'BRKM5.SA', 'BRFS3.SA', 'BPAC11.SA', 'CRFB3.SA', 'CCRO3.SA', 'CMIG4.SA', 'CIEL3.SA', 
            'COGN3.SA', 'CPLE6.SA', 'CSAN3.SA', 'CPFE3.SA', 'CMIN3.SA', 'CVCB3.SA', 'CYRE3.SA', 
            'DXCO3.SA', 'ECOR3.SA', 'ELET3.SA', 'ELET6.SA', 'EMBR3.SA', 'ENBR3.SA', 'ENGI11.SA', 
            'ENEV3.SA', 'EGIE3.SA', 'EQTL3.SA', 'EZTC3.SA', 'FLRY3.SA', 'GGBR4.SA', 'GOAU4.SA', 
            'GOLL4.SA', 'NTCO3.SA', 'SOMA3.SA', 'HAPV3.SA', 'HYPE3.SA', 'IGTI11.SA', 'ITSA4.SA', 
            'ITUB4.SA', 'JBSS3.SA', 'KLBN11.SA', 'RENT3.SA', 'LWSA3.SA', 'LREN3.SA', 'MGLU3.SA', 
            'MRFG3.SA', 'CASH3.SA', 'BEEF3.SA', 'MRVE3.SA', 'MULT3.SA', 'PCAR3.SA', 'PETR3.SA', 
            'PETR4.SA', 'PRIO3.SA', 'PETZ3.SA', 'QUAL3.SA', 'RADL3.SA', 'RAIZ4.SA', 'RDOR3.SA', 
            'RAIL3.SA', 'SBSP3.SA', 'SANB11.SA', 'SMTO3.SA', 'CSNA3.SA', 'SLCE3.SA', 'SUZB3.SA', 
            'TAEE11.SA', 'VIVT3.SA', 'TIMS3.SA', 'TOTS3.SA', 'UGPA3.SA', 'USIM5.SA', 'VALE3.SA', 
            'VIIA3.SA', 'VBBR3.SA', 'WEGE3.SA', 'YDUQ3.SA']

# Busca as açoes usando o Yfinance e os armazena em um DataFrame
def busca_acoes(lista):
    data = datetime.now()
    nova_lista = lista
    dados_acoes = yf.download(nova_lista, period='1d', start=ultimo_dia_util_anterior(data))
    closing_prices = dados_acoes['Close']
    fech = closing_prices.transpose().fillna(0)
    fech.loc[:, "Código"] = nova_lista
    return fech
    
    
# Busca qual foi o ultimo dia útil anterior
def ultimo_dia_util_anterior(data:datetime):
    # Verifica se a data fornecida é um dia útil

    def is_dia_util(data:datetime):
        if data.weekday() < 5:  # Considera dias de segunda a sexta-feira como dias úteis
            return data
        else:
            return None

    # Retrocede um dia
    data_anterior = is_dia_util(data) - timedelta(days=1)
    
    # Se a data anterior for um dia útil, retorna essa data
    if is_dia_util(data_anterior):
        return data_anterior
    
    # Se não for um dia útil, retrocede até encontrar o último dia útil
    while not is_dia_util(data_anterior):
        data_anterior -= timedelta(days=1)
    return data_anterior

# Pega o resultado da funcao busca_acoes() e abre uma janela para o usuário salvar o arquivo
def salvar_excel(tabela):
    root = Tk()
    root.withdraw()  # Esconde a janela principal
    # Permite ao usuário escolher o local e o nome do arquivo
    nome_arquivo = filedialog.asksaveasfilename(
    defaultextension=".xlsx",
    filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
    initialfile=f"tabela_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
    title="Salvar Tabela"
)

    # Verifica se o usuário cancelou a seleção
    if not nome_arquivo:
        print("Seleção de arquivo cancelada.")
    else:
    # Salva a tabela no arquivo Excel
        tabela.to_excel(nome_arquivo, index=False)
        print(f"Tabela salva com sucesso em {nome_arquivo}")

def clique():
    df = busca_acoes(acoes)
    salvar_excel(df)

        
def janela_busca():  
    sg.theme('Reddit')
    layout = [
        [sg.Text('BUSCADOR DE COTAÇÕES DA B3')],
        [sg.Text('Para baixar: ')],[sg.Button('Baixar')], 
       [sg.Text('')],[sg.Button('Fechar')]]
    
    return sg.Window('Buscador de Cotação B3', layout=layout, finalize=True)

janela1 = janela_busca()

while True:
    window, event, values = sg.read_all_windows()
    if window == janela1 and event == sg.WIN_CLOSED or event == 'Fechar':
        break
    if window == janela1 and event == 'Baixar':
        clique()
        janela1.hide()
    


