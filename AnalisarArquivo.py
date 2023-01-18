import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)

# Tabela de faturamento por loja
tabela_loja = tabela_vendas[['ID Loja','Valor Final']]\
    .groupby('ID Loja').sum().sort_values(by='Valor Final', ascending=False)
print(tabela_loja)
print('-' * 50)

# Tabela de qtd vendida por loja
tabela_qtd_vendida = tabela_vendas[['ID Loja','Quantidade']]\
    .groupby('ID Loja').sum().sort_values(by='Quantidade', ascending=False)
print(tabela_qtd_vendida)
print('-' * 50)

# Ticket médio por loja
tabela_vendas['Ticket Médio'] = tabela_vendas['Valor Final'] / tabela_vendas['Quantidade']
tabela_ticket_medio = tabela_vendas[['ID Loja', 'Ticket Médio']]\
    .groupby('ID Loja').mean().sort_values(by='Ticket Médio', ascending=False)
print(tabela_ticket_medio)

# Enviar e-mail

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'sopradivertir1@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue abaixo relatórios de vendas.</p>

<p>Faturamento por loja:</p>
{tabela_loja.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida por loja:</p>
{tabela_qtd_vendida.to_html()}

<p>Ticket Médio por loja:</p>
{tabela_ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Atenciosamente,</p>
<p>Matheus Araujo Morais</p>
'''
mail.Send()
print("\nE-mail enviado")
