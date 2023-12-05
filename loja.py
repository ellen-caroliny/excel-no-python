import pandas as pd
import win32com.client as win32

tabelaVendas = pd.read_excel('./Vendas.xlsx')

pd.set_option('display.max_columns', None)
print(tabelaVendas)

faturamento = tabelaVendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

qtde= tabelaVendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(qtde)
print('-'* 50)

ticketMedio = (faturamento['Valor Final']/ qtde['Quantidade']).to_frame()
ticketMedio - ticketMedio.rename(columns={0: 'Ticket Médio'})
print(ticketMedio)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'ellensouza4666@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{qtde.to_html()}

<p>Ticket médio do produto por cada loja:</p>
{ticketMedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p> Entre em contato para quaisquer dúvidas</p>


'''
mail.Send()
print('E-mail enviado')