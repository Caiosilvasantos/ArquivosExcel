#importar base de dados
import pandas as pd
import win32com.client as win32

#visualizar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns',None)
print(tabela_vendas)
#faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)
#ticket medio por produtos vendidos por loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()#para deixar como tabela
ticket_medio = ticket_medio.rename(columns={0:'Ticket medio'})
print(ticket_medio)

#enviar email com relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email.com'
mail.Subject = 'Relatorios vendas por cada Loja'
mail.Body = 'Message body'
mail.HTMLBody = f''' 

<p> Prezados, </p>
<p>segue o relatorio de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>quantidade Vendida:</p>
{quantidade.to_html()}

<p> Ticket medio dos produtos de cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Att</p>
<p>Fulano</p>

'''
# To attach a file to the email (optional):

mail.Send()
print("-----------------------------------")
print("Enviado com Sucesso!!")