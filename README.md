# Processamento e Visualização de Dados de Funcionários
Este projeto tem como objetivo processar e visualizar dados de funcionários a partir de vários arquivos Excel. O foco está em calcular o quantitativo de funcionários por empresa, mês e ano, garantindo que cada funcionário seja contabilizado apenas uma vez por mês.

## Requisitos
Antes de rodar o script, certifique-se de que você tem as seguintes bibliotecas instaladas:

- ** pandas ** -
- ** matplotlib ** -
- ** openpyxl ** -

## Instalação
Para instalar as dependências, execute o seguinte comando:

`pip install pandas matplotlib openpyxl`
# Estrutura do Projeto
Diretório de Dados: Os arquivos Excel contendo os dados estão localizados em um diretório especificado no código.
Colunas de Interesse: Empresa, Matrícula, Nome, e Data Referência são as colunas usadas para o processamento.

# Funcionalidades

## O script faz o seguinte:
Carrega múltiplos arquivos Excel de um diretório especificado.
Extrai as colunas relevantes dos arquivos Excel: Empresa, Matrícula, Nome, e Data Referência.
Concatena os arquivos em um único DataFrame, unindo os dados de todas as planilhas.
Cria um identificador único para cada funcionário, concatenando Matrícula e Nome.
Remove duplicatas: Cada funcionário é contabilizado apenas uma vez por mês e por empresa.
Agrupa os dados por empresa, mês e ano para calcular o total de funcionários únicos em cada período.
Gera um gráfico mostrando o período (ano-mês) que teve o maior número de funcionários.
Exibe o quantitativo de funcionários por empresa, mês e ano no console.

# Estrutura do Código
## 1. Importação de Bibliotecas

`import pandas as pd`
`import os`
`import glob`
`import matplotlib.pyplot as plt`

## 2. Carregamento dos Dados
Os arquivos são carregados de um diretório específico:

`caminho_diretorio = r'\\172.21.48.102\d$\dev\dataBaseBI\Folha - Backup'`
`dataframes = []`

Para cada arquivo Excel no diretório, o código extrai as colunas Empresa, Matrícula, Nome, e Data Referência:

`for caminho_arquivo in glob.glob(os.path.join(caminho_diretorio, '*.xlsx')):`
    `try:`
        `dados = pd.read_excel(caminho_arquivo, usecols=[`
            `'Empresa', 'Matrícula', 'Nome', 'Data Referência'`
        `])`
        `dataframes.append(dados)`
    `except Exception as e:`
        `print(f"Erro ao processar o arquivo {caminho_arquivo}: {e}")`
        
## 3. Concatenação e Criação de Identificador

Os dados de todos os arquivos são concatenados e é criado um identificador único para cada funcionário, combinando Matrícula e Nome:`

`dados_concatenados = pd.concat(dataframes, ignore_index=True)`
`dados_concatenados['Identificador'] = dados_concatenados['Matrícula'].astype(str) + '-' + dados_concatenados['Nome']`

## 4. Conversão da Data e Remoção de Duplicatas

A coluna Data Referência é convertida para um formato de data. Além disso, criamos colunas separadas para o Ano e o Mês:

`dados_concatenados['Data Referência'] = pd.to_datetime(dados_concatenados['Data Referência'], dayfirst=True, errors='coerce')`
`dados_concatenados['Ano'] = dados_concatenados['Data Referência'].dt.year`
`dados_concatenados['Mês'] = dados_concatenados['Data Referência'].dt.month`

## Removemos duplicatas para garantir que cada funcionário seja contado apenas uma vez por mês:

`dados_unicos = dados_concatenados.drop_duplicates(subset=['Identificador', 'Ano', 'Mês'])`

## 5. Agrupamento e Cálculo de Funcionários
Agrupamos os dados por Empresa, Ano e Mês para calcular o total de funcionários únicos:

quantitativo_funcionarios = dados_unicos.groupby(['Empresa', 'Ano', 'Mês']).size().reset_index(name='Total Funcionários')
## 6. Visualização Gráfica
Geramos um gráfico mostrando o período com o maior número de funcionários:

periodo_maximo = quantitativo_funcionarios.groupby(['Ano', 'Mês'])['Total Funcionários'].sum().reset_index()
periodo_maximo_sorted = periodo_maximo.sort_values(by='Total Funcionários', ascending=False)

plt.figure(figsize=(10,6))
plt.plot(periodo_maximo_sorted['Ano'].astype(str) + '-' + periodo_maximo_sorted['Mês'].astype(str), periodo_maximo_sorted['Total Funcionários'], marker='o')
plt.title('Período com mais funcionários')
plt.xlabel('Período (Ano-Mês)')
plt.ylabel('Total de Funcionários')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

## 7. Exibição dos Resultados
Por fim, exibimos o quantitativo de funcionários processado:

`print(quantitativo_funcionarios)`

## Como Usar
Coloque os arquivos Excel no diretório especificado.
Execute o script para gerar o gráfico e ver os dados processados.
Os dados processados serão impressos no console, e o gráfico exibirá o período com o maior número de funcionários.

## Exemplo de Saída

`Empresa      Ano  Mês  Total Funcionários`
`Empresa A    2023   4          1700`
`Empresa B    2024   8          1713`

O gráfico exibirá o período com o maior número de funcionários, organizando por ano e mês.
