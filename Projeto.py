import pandas as pd, openpyxl
import win32com.client as win32

#Pandas será ultilizado para ler os arquivos e
#Openpyxl specializada para trabalhar com arquivos Excel

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')
#tabela_vendas recebe via biblioteca a leitura da base de dados em excel
print('-'*50)

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
print('-'*50)

# Ticket Médio por produto em cada loja
ticket_medio = (tabela_faturamento['Valor Final'] / tabela_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
#Mudando o nome da coluna
#Dividindo uma tabela de bados pela outra para obter a media e armazenando no ticket_medio
#.to_frame() serve para formar em uma tabela os dados divididos
print(ticket_medio)

# Enviar um E-mail com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.To = 'romariodiogo2021@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''

<p>Prezados,</p>
<p>Segue o Relatório de Vendass por cada Loja.</p>

<p>Faturamento:</p>
{tabela_faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p> 
{tabela_produtos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja: </p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>


<p>Atte.,</p>
<p>Romário </p>


'''

mail.Send()

print('E-mail Enviado')