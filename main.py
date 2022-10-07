import pandas as pd
import win32com.client as win32
# Importação Base de Dados
tabela_vendas = pd.read_excel('vendas.xlsx')

# Visualização Base de Dados
print('=================== Base de Dados ===================')
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja
print('=================== Faturamento ===================')
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de Produtos Vendidos por loja
print('=================== Produtos Vendidos ===================')
quant = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quant)

# Ticket médio por produto em cada loja (Faturamento/Quantidade)
print('=================== Ticket Médio ===================')
ticket_med = (faturamento['Valor Final'] / quant['Quantidade']).to_frame()
ticket_med = ticket_med.rename(columns={0: 'Ticket Médio'})
print(ticket_med)

# Enviar um e-mail com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'joaofelipecrodrigues@gmail.com' # Endereço para quem ira o e-mail
mail.Subject = 'Relatório de Vendas por Loja' # Assunto do e-mail
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quant.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_med.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>João F. C. Rodrigues</p>
'''

mail.Send()
print('Email Enviado')