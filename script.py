import win32com.client as win32

try:
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'tiburciomartim@gmail.com'
    mail.Subject = 'Relatório de vendas por loja'
    mail.HTMLBody = '''
        Prezados,

        Segue o relatório de vendas por cada loja.
        Faturamento:
        {}

        Quantidade Vendida:
        {}

        Ticket Médio dos Produtos em cada Loja:
        {}

        Qualquer dúvida estou à disposição.

        Att.,
        Tiburcio
    '''
    print('Enviado com sucesso!')
    mail.Send()
except Exception as e:
    print(f"Falha na conexão. {e}")