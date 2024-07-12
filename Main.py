import pandas as pd
import win32com.client as win32
#Primeiro passo: Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


#Segundo passo: Visuaçlizar a base de daos 
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# #Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#  Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# ticket médio por produto em cada loja 
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'mouracamily93 @gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p><strong>Faturamento:</strong></p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><strong>Quantidade Vendida:</strong></p>
{quantidade.to_html()}

<p><strong>Ticket Médio dos Produtos em cada Loja:</strong></p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</strong></p>

<p>Att.,</p>
<p>Wenderson</p>

'''

mail.Send()

print('Email Enviado')