import pandas as pd, openpyxl
#Pandas será ultilizado para ler os arquivos e
#Openpyxl specializada para trabalhar com arquivos Excel

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')
#tabela_vendas recebe via biblioteca a leitura da base de dados em excel

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
#Mostrando todas as colunas
print(tabela_vendas)


# Faturamento
tabela_faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#[['ID Loja', 'Valor Final']] -> Filtrando as colunas que desejo
#.groupby('ID Loja') -> Agrupando todas as lojas
#.sum() -> Somando todas as lojas

#print(tabela_faturamento)

# Quantidade de produtos vendidos por loja
tabela_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print(tabela_produtos)

# Ticket Médio por produto em cada loja

# Enviar um E-mail com o relatório

