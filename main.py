                                    # pip install openpyxl    | Read files in Excel
import pandas as pd                 # pip install pandas      | To work with tables
import win32com.client as win32     # pip install pywin32     | Integrating Python with windows

# To import data base
tabela_vendas = pd.read_excel("Vendas.xlsx")

# To view data base
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)

# Billing per store
earnings = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(earnings)
print('-' * 50)

# Quantity of products sell per store
quantity = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantity)
print('-' * 50)

# Averige ticket per products in each store
ticket_medio = (earnings['Valor Final'] / quantity['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Tiket Médio'})
print(f'\n\n{ticket_medio}')
print('-' * 50)

try:
    # Send email with report
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'tiburciomartim@gmail.com'
    mail.Subject = 'Relatório de vendas por loja'
    mail.HTMLBody = f'''
        <p>Prezados,</p>
        
        <p>Segue o relatório de vendas por cada loja.</p>
        <p>Faturamento:</p>
        {earnings.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}
        
        <p>Quantidade Vendida:</p>
        {quantity.to_html()}
        
        <p>Ticket Médio dos Produtos em cada Loja:</p>
        {ticket_medio.to_html(formatters={'Tiket Médio': 'R$ {:,.2f}'.format})}
        
        <p>Qualquer dúvida estou à disposição.</p>
        
        <p>Att.,</p>
        <p>Tiburcio</p>
    '''
    mail.Send()
    print('\nEmail enviado.')

except Exception as e:
    print(f"Falha na conexão. {e}")

