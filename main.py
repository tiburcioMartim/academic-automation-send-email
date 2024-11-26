# Necessário para ler arquivos em Excel: # pip install openpyxl
import pandas as pd             # pip install pandas
import win32com.client as win32    # pip install pywin32 | # Integra o python com o windows:


# importar base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")


# visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)


# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)


# quantidade de produtos vendidos por loja
quantidades = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidades)
print('-' * 50)


# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidades['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Tiket Médio'})
print(f'\n\n{ticket_medio}')
print('-' * 50)


try:
    # enviar email com o relatório
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'tiburciomartim@gmail.com'
    mail.Subject = 'Relatório de vendas por loja'
    mail.HTMLBody = f'''
        <p>Prezados,</p>
        
        <p>Segue o relatório de vendas por cada loja.</p>
        <p>Faturamento:</p>
        {faturamento.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}
        
        <p>Quantidade Vendida:</p>
        {quantidades.to_html()}
        
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

