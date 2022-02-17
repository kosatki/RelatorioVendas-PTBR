# 1: instalação e importação de bibliotecas
import pandas as pd
import openpyxl

# Importação da base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualização da base de dados
pd.set_option('display.max_columns', None)


# Agrupando por lojas
# tabela_vendas.groupby('ID Loja').sum()

# Agrupando por lojas, filtrando e somando colunas
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)

# SOmando a quantidade de itens vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)

# Calculando a média em reais por item vendido por loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)



# INTEGRAÇÃO DO EMAIL COM O PYTHON USANDO PYWIN32

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email para envio'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p> 

<p>Segue relatório de vendas por loja.</p> 

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio por Produto em cada Loja:</p>
{ticket_medio.to_html(formatters = {'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att,</p>
<p>Kosatki.</p>
'''

mail.Send()

print('E-mail enviado com sucesso.')