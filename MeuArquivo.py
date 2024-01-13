import pandas as pd
import openpyxl
import win32com.client as win32


# importar base de dados -> pandas faz a iteração entre python e excel
tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados para possíveis correções e tirar ocultação das colunas
pd.set_option('display.max_columns', None) 
#print(tabela_vendas)

# Calcular faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' *50)
print('-' *50)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' *50)
print('-' *50)

# ticket médio por produto em cada loja -> Calcular faturamento / qtd vendida
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() 
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)
print('-' *50)
print('-' *50)

# enviar e-mail com relatório -> importar biblioteca pywin32 pra integrar com outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'graziellacoutinho@yahoo.com.br'
mail.Subject = 'Relatório'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue Relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Atenciosamente,</p>
<p>Graziella Rodrigues</p>

 
'''
mail.Send()
print('Email-enviado!')

