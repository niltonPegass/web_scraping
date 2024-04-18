from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import time
import os
import requests
import pandas as pd
# import tkinter as tk
# from tkinter import filedialog

# caminho_string = r'C:\Users\Documentos'
caminho_string = str(input("informe o caminho que deseja salvar os arquivos: "))
inicio = str(input("informe a data inicial [dd/mm/aaaa]: "))
final = str(input("informe a data final [dd/mm/aaaa]: "))
intervalo_requisicoes = int(input("informe o intervalo entre as requisições [em segundos]: "))

dia_inicio, mes_inicio, ano_inicio = map(int, inicio.split('/'))
dia_final, mes_final, ano_final = map(int, final.split('/'))

# função para extrair os dados da tabela
def extrair_dados(html):
    soup = BeautifulSoup(html, 'html.parser')
    tabela = soup.find('table', {'id': 'example'})
    linhas = tabela.find_all('tr')
    
    dados = []
    for linha in linhas[1:]:  # ignora o cabeçalho
        colunas = linha.find_all('td')
        matricula = colunas[0].text.strip()
        nome = colunas[1].text.strip()
        cargo = colunas[2].text.strip()
        empresa = colunas[3].text.strip()
        data = colunas[4].text.strip()
        destino = colunas[5].text.strip()
        motivo = colunas[6].text.strip()
        diarias = colunas[7].text.strip()
        valor = colunas[8].text.strip()
        dados.append((matricula, nome, cargo, empresa, data, destino, motivo, diarias, valor))
    
    return dados

# função para realizar a solicitação HTTP e extrair os dados
def scrap_data(data):
    url = 'http://www2.voltaredonda.rj.gov.br/transparencia/mod/guia-de-viagem/'
    payload = {'data': data, 'enviar': 'enviar'}
    headers = {'User-Agent': 'Mozilla/5.0'}

    response = requests.post(url, data=payload, headers=headers)
    if response.status_code == 200:
        return extrair_dados(response.text)
    else:
        print('erro ao acessar a página:', response.status_code)
        return None

# função para converter uma string de data para o formato 'dd/mm/yyyy'
def converter_data(data):
    return datetime.strptime(data, '%d/%m/%Y')

# Exemplo de uso
data_inicial = datetime(ano_inicio, mes_inicio, dia_inicio)  # data inicial
data_final = datetime(ano_final, mes_final, dia_final)   # data final

dados_totais = []
data_atual = data_inicial
while data_atual <= data_final:
    data_formatada = converter_data(data_atual.strftime('%d/%m/%Y'))
    print(f'Obtendo dados para {data_formatada.date()}')
    dados = scrap_data(data_formatada.strftime('%d/%m/%Y'))
    if dados:
        dados_totais.extend(dados)

    # contagem regressiva até a próxima requisição
    if data_atual < data_final:
        for segundos_restantes in range(intervalo_requisicoes, 0, -1):
            print(f'tempo restante para próxima requisição: {segundos_restantes} segundos', end='\r')
            time.sleep(1)
        print("")
    
    data_atual += timedelta(days=1)

if dados_totais:
    df = pd.DataFrame(dados_totais, columns=['Matrícula', 'Nome', 'Cargo', 'Empresa', 'Data', 'Destino', 'Motivo', 'Diárias', 'Valor'])
    
    # root = tk.Tk()
    # root.withdraw()  # Esconde a janela principal
    # pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar os dados")
    ##pasta_destino = os.path.join('I:\\', 'Nilton', 'tech studies', 'DIO')

    def criar_caminho_pasta(caminho_string):
        partes = caminho_string.split(os.sep)
        identificador_disco = partes[0] + '\\'

        # contando o número de partes
        # num_pastas = len(partes)
        
        pasta_destino = os.path.join(identificador_disco, *partes)
    
        return pasta_destino #, num_pastas

    pasta_destino = criar_caminho_pasta(caminho_string)

    if pasta_destino:
        for mes_ano, df_mes in df.groupby(df['Data'].apply(lambda x: datetime.strptime(x, '%d/%m/%Y').strftime('%Y.%m'))):
            nome_arquivo = f'dados_viagens_{mes_ano}.xlsx'
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
            df_mes.to_excel(caminho_arquivo, index=False)
            print(f'dados exportados para {caminho_arquivo} com sucesso')
            
    else:
        print('nenhum diretório selecionado. os dados não foram exportados')
else:
    print('não foi possível obter os dados')
