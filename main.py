import win32com.client as win32
import pandas as pd
import os

caminho = os.getcwd() # pega o caminho do diretório atual

tabela = pd.read_excel('base_clientes.xlsx')
tabela['Valor Total'] = tabela['Quantidade'] * tabela['Valor Unitário']
tabela['Email'] = tabela['Email'].str.lower().str.replace(' ', '') # Limpeza da coluna de email

clientes = tabela['Clientes'].unique()
outlook = win32.Dispatch('outlook.application')

for cliente in clientes:
    tabela_cliente = tabela[tabela["Clientes"] == cliente]  #Filtra da tabela original os Clientes
    valor_venda = tabela_cliente["Valor Total"].sum()       #trata o Valor total da venda para mandar no email

    nome_do_arquivo = f'{cliente}.xlsx'
    tabela_cliente.to_excel(nome_do_arquivo, index=False) #Cria o arquivo com o nome da empresa
    
    email = outlook.CreateItem(0)
    email.To = tabela_cliente['Email'].iloc[0] 
    email.Subject = f'Relatório do {cliente}'
    email.Body = f"""Prezado {cliente},

Segue em anexo o relatório de vendas. O valor total das suas compras foi de R${valor_venda:.2f}.

Atenciosamente,
Teste"""
    
    email.Attachments.Add(os.path.join(caminho, nome_do_arquivo))
    email.Send()  #se usar display() ele abre o outlook com o email, destinatário e anexo preenchidos mas não envia para mostrar se está funcionando