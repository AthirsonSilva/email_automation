import pandas as pd
import win32com.client as win32

# importing the database
sales_table = pd.read_excel('./databases/Vendas.xlsx')

# viewing the database
pd.set_option('display.max_columns', None)
print(sales_table)

# profit per store
profit = sales_table[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(profit)

# quantity of products saled per store
quantity = sales_table[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantity)

print('-' * 50)

# Average ticket per product on each store
average_ticket = (profit['Valor Final'] / quantity['Quantidade']).to_frame()
average_ticket = average_ticket.rename(columsn = {0: 'Ticket Médio'})

# Send the email with a relatory
outlook = wi32.Dispatch('outlook.application')
mail = outloo.CreateItem(0)
mail.To = 'Athirsonarceus@gmail.com'
mail.Subject = 'Sales per store relatory'
mail.HTMLBody = f"""
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{profit.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantity.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{average_ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Athirson</p>
"""

mail.Send()

print('Email sent!')