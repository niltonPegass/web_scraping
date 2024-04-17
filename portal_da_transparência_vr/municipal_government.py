from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from tkinter import filedialog
import tkinter as tk
import os
import time
import requests
import pandas as pd

# caminho_string = r'P:\Comercial\Estudo Conveniência'
inicio = str(input("informe a data inicial [dd/mm/aaaa]: "))
final = str(input("informe a data final [dd/mm/aaaa]: "))
intervalo_requisicoes = int(input("informe o intervalo entre as requisições [seg]: "))

dia_inicio, mes_inicio, ano_inicio = map(int, inicio.split('/'))
dia_final, mes_final, ano_final = map(int, final.split('/'))

# Função para extrair os dados da tabela
def extrair_dados(html):
    soup = BeautifulSoup(html, 'html.parser')
    tabela = soup.find('table', {'id': 'example'})
    linhas = tabela.find_all('tr')
    
    dados = []
    for linha in linhas[1:]:  # Ignora o cabeçalho
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

# Função para realizar a solicitação HTTP e extrair os dados
def scrap_data(data):
    url = 'http://www2.voltaredonda.rj.gov.br/transparencia/mod/guia-de-viagem/'
    payload = {'data': data, 'enviar': 'enviar'}
    headers = {'User-Agent': 'Mozilla/5.0'}

    response = requests.post(url, data=payload, headers=headers)
    if response.status_code == 200:
        return extrair_dados(response.text)
    else:
        print('Erro ao acessar a página:', response.status_code)
        return None

# Função para converter uma string de data para o formato 'dd/mm/yyyy'
def converter_data(data):
    return datetime.strptime(data, '%d/%m/%Y')

# Exemplo de uso
data_inicial = datetime(ano_inicio, mes_inicio, dia_inicio)  # Data inicial
data_final = datetime(ano_final, mes_final, dia_final)   # Data final

dados_totais = []
data_atual = data_inicial
while data_atual <= data_final:
    data_formatada = converter_data(data_atual.strftime('%d/%m/%Y'))
    print(f'Obtendo dados para {data_formatada.date()}')
    dados = scrap_data(data_formatada.strftime('%d/%m/%Y'))
    if dados:
        dados_totais.extend(dados)
    data_atual += timedelta(days=1)
    
    # Contagem regressiva até a próxima requisição
    for segundos_restantes in range(intervalo_requisicoes, 0, -1):
        print(f'Tempo restante para próxima requisição: {segundos_restantes} segundos', end='\r')
        time.sleep(1)
    print("")

if dados_totais:
    # Convertendo os dados para um DataFrame do pandas
    df = pd.DataFrame(dados_totais, columns=['Matrícula', 'Nome', 'Cargo', 'Empresa', 'Data', 'Destino', 'Motivo', 'Diárias', 'Valor'])
    
    # Criando uma pasta para salvar os arquivos
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar os dados")
    # pasta_destino = os.path.join('I:\\', 'Nilton', 'tech studies', 'DIO')

    # def criar_caminho_pasta(caminho_string):
    #     # Dividindo o caminho em partes
    #     partes = caminho_string.split(os.sep)
    #     identificador_disco = partes[0] + '\\'

    #     # Contando o número de partes
    #     num_pastas = len(partes)
        
    #     # Construindo o caminho usando os.path.join()
    #     pasta_destino = os.path.join(identificador_disco, *partes)
    
    #     return num_pastas, pasta_destino

    # pasta_destino = criar_caminho_pasta(caminho_string)

    if pasta_destino:  # Se um diretório foi selecionado
        for mes_ano, df_mes in df.groupby(df['Data'].apply(lambda x: datetime.strptime(x, '%d/%m/%Y').strftime('%Y.%m'))):
            # Criando o nome do arquivo com base no mês e ano
            nome_arquivo = f'dados_viagens_{mes_ano}.xlsx'
            # Criando o caminho completo do arquivo dentro da pasta de destino
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
            # Exportando para o arquivo Excel
            df_mes.to_excel(caminho_arquivo, index=False)
            print(f'Dados exportados para {caminho_arquivo} com sucesso.')
    else:
        print('Nenhum diretório selecionado. Os dados não foram exportados.')
else:
    print('Não foi possível obter os dados.')
