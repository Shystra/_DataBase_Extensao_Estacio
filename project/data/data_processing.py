import pandas as pd
import os
import glob
import matplotlib.pyplot as plt

caminho_diretorio = r'\\172.21.48.102\d$\dev\dataBaseBI\Folha'
dataframes = []

for caminho_arquivo in glob.glob(os.path.join(caminho_diretorio, '*.xlsx')):
    try:
        dados = pd.read_excel(caminho_arquivo, usecols=[
            'Empresa', 'Matrícula', 'Nome', 'Data Referência'
        ])
        dataframes.append(dados)
    except Exception as e:
        print(f"Erro ao processar o arquivo {caminho_arquivo}: {e}")

dados_concatenados = pd.concat(dataframes, ignore_index=True)
dados_concatenados['Identificador'] = dados_concatenados['Matrícula'].astype(str) + '-' + dados_concatenados['Nome']
dados_concatenados['Data Referência'] = pd.to_datetime(dados_concatenados['Data Referência'], dayfirst=True, errors='coerce')
dados_concatenados['Ano'] = dados_concatenados['Data Referência'].dt.year
dados_concatenados['Mês'] = dados_concatenados['Data Referência'].dt.month

dados_unicos = dados_concatenados.drop_duplicates(subset=['Identificador', 'Ano', 'Mês'])
quantitativo_funcionarios = dados_unicos.groupby(['Empresa', 'Ano', 'Mês']).size().reset_index(name='Total Funcionários')

periodo_maximo = quantitativo_funcionarios.groupby(['Ano', 'Mês'])['Total Funcionários'].sum().reset_index()
periodo_maximo_sorted = periodo_maximo.sort_values(by='Total Funcionários', ascending=False)

print(quantitativo_funcionarios)

plt.figure(figsize=(10,6))
plt.plot(periodo_maximo_sorted['Ano'].astype(str) + '-' + periodo_maximo_sorted['Mês'].astype(str), periodo_maximo_sorted['Total Funcionários'], marker='o')
plt.title('Período com mais funcionários')
plt.xlabel('Período (Ano-Mês)')
plt.ylabel('Total de Funcionários')
plt.xticks(rotation=45)
plt.tight_layout()

plt.show()

