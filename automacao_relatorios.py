
import pandas as pd
import win32com.client as win32
import pathlib


vendas_df=pd.read_excel(r'Bases de Dados\Vendas.xlsx')
lojas_df=pd.read_csv(r'Bases de Dados\Lojas.csv',encoding='latin1',sep=';')
emails_df=pd.read_excel(r'Bases de Dados\Emails.xlsx')

print(vendas_df)
print(lojas_df)
print(emails_df)


vendas_df=vendas_df.merge(lojas_df,on='ID Loja')
print(vendas_df)

dicionario_lojas={}
for loja in lojas_df['Loja']:
    dicionario_lojas[loja] = vendas_df.loc[vendas_df['Loja']==loja, :]

print(dicionario_lojas)

dia_indicador = vendas_df['Data'].max()
print(f' O dia indicador é {dia_indicador.day}/{dia_indicador.month}')


caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pastas_backup = caminho_backup.iterdir()

lista_arquivos_backup=[arquivo.name for arquivo in arquivos_pastas_backup]
print(lista_arquivos_backup)

for loja in dicionario_lojas:
    try:
        if loja not in lista_arquivos_backup:
            nova_pasta = caminho_backup / loja
            nova_pasta.mkdir()
    except FileExistsError:
        pass
    
    nome_arquivo='{}_{}_{}.xlsx'.format(dia_indicador.month,dia_indicador.day,loja)
    local_arquivo = caminho_backup/loja/nome_arquivo
            
    dicionario_lojas[loja].to_excel(local_arquivo)

    meta_faturamento = 1000
meta_faturamento_ano=1650000
meta_qtdprodutos_dia = 4
meta_qtdprodutos_ano = 120
meta_ticketmedio_dia= 500
meta_ticketmedio_ano = 500


for loja in dicionario_lojas:
    #criando dicionario para as pegar o dia indicador da loja e o total de vendas
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador,:]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
   


    #Quantidade de Itens
    qtd_produtos_loja= len(vendas_loja['Produto'].unique())
    qtd_produtos_loja_dia=len(vendas_loja_dia['Produto'])
    



    #Ticket Medio
    valor_venda=vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    valor_vendas_dia=vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_vendas_dia['Valor Final'].mean()
    


    #enviando email
    outlook = win32.Dispatch('outlook.application')
    nome = emails_df.loc[emails_df['Loja'] == loja, 'Gerente'].values[0] #nome do gerente da loja
    email = outlook.CreateItem(0)
    email.To = 'mrcartonrex@gmail.com' #emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    email.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}' #assunto

    cor_maior = 'green' #cor verde
    cor_menor = 'red' #cor vermelha

    #Faturamento condição
    if faturamento_dia>=meta_faturamento:
        dia_fatu = cor_maior
    else:
        dia_fatu = cor_menor

    if faturamento_ano>=meta_faturamento_ano:
          ano_fat = cor_maior
    else:
        ano_fat = cor_menor

    #Quantidade Produtos condição
    if qtd_produtos_loja_dia>=meta_qtdprodutos_dia:
        dia_qtd = cor_maior
    else:
        dia_qtd = cor_menor

    if qtd_produtos_loja>=meta_qtdprodutos_ano:
          ano_qtd = cor_maior
    else:
        ano_qtd = cor_menor

    #Ticket Medio condição
    if ticket_medio_dia>=meta_ticketmedio_dia:
          dia_ticket = cor_maior
    else:
        dia_ticket = cor_menor
    if ticket_medio_ano>=meta_ticketmedio_ano:
          ano_ticket = cor_maior
    else:
        ano_ticket = cor_menor

    #texto do email
    email.HTMLBody = f""" 
    <p>Bom dia {nome}</p>
    <p>O resultado do dia <strong>{dia_indicador.day}/{dia_indicador.month}</strong> da <strong>Loja {loja}</strong> foi de: </p>

    <table>
      <tr>
        <th></th>
        <th style="text-align: center">Valor Dia</th>
        <th style="text-align: center">Meta Dia </th>
        <th style="text-align: center">Cénario Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R$ {faturamento_dia:.2f}</td>
        <td style="text-align: center"> R$ {meta_faturamento:.2f}</td>
        <td style="text-align: center"><font color ={dia_fatu}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtd_produtos_loja_dia}</td>
        <td style="text-align: center">{meta_qtdprodutos_dia}</td>
        <td style="text-align: center"><font color ={dia_qtd}>◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio por Venda</td>
        <td style="text-align: center">R$ {ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R$ {meta_ticketmedio_dia}</td>
        <td style="text-align: center"><font color ={dia_ticket}>◙</font></td>
      </tr>
      </table>
      <br>
      <table>
      <tr>
        <th></th>
        <th style="text-align: center">Valor Ano</th>
        <th style="text-align: center">Meta Ano </th>
        <th style="text-align: center">Cénario Ano</th>
      </tr>
      <tr>
        <td >Faturamento</td>
        <td style="text-align: center">R$ {faturamento_ano:.2f}</td>
        <td style="text-align: center">R$ {meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color ={ano_fat}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade De Produtos</td>
        <td style="text-align: center">{qtd_produtos_loja}</td>
        <td style="text-align: center">{meta_qtdprodutos_ano}</td>
        <td style="text-align: center"><font color ={ano_qtd}>◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio por Venda</td>
        <td style="text-align: center">R$ {ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R$ {meta_ticketmedio_ano}</td>
        <td style="text-align: center"><font color ={ano_ticket}>◙</font></td>
      </tr>
    </table>


    <p> Segue um anexo da planilha para mais detalhes.</p>

    <p>Qualquer dúvida estou a disposição </p>
    <p>Att., Jader</p>

    """
    #buscando documento dentro da pasta do PC e definindo padrao para enviar por email
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    email.Attachments.Add(str(attachment)) #transformando o caminho em texto para o python entender onde procurar


    email.Send() #email enviado
    print(f"O e-mail da loja {loja} foi enviado")

    faturamento_lojas = vendas_df.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending = False)


vendas_dia = vendas_df.loc[vendas_df['Data']==dia_indicador,:]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending = False)



nome_arquivo_ano='{}_{}_Ranking_Anual.xlsx'.format(dia_indicador.month,dia_indicador.day)          
faturamento_lojas_ano.to_excel(r'{}\{}'.format(caminho_backup,nome_arquivo_ano))
nome_arquivo_dia='{}_{}_Ranking_Dia.xlsx'.format(dia_indicador.month,dia_indicador.day) 
faturamento_lojas_dia.to_excel(r'{}/{}'.format(caminho_backup,nome_arquivo_dia))


import win32com.client as win32
import pathlib
#enviando email
outlook = win32.Dispatch('outlook.application')
nome = emails_df.loc[emails_df['Loja'] ==  'Diretoria', 'E-mail'].values[0] 
mail = outlook.CreateItem(0)
mail.To = 'mrcartonrex@gmail.com' #emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]email.Subject = f' Ranking Dia {dia_indicador.day}/{dia_indicador.month} ' #assunto
mail.Subject = f'Ranking {dia_indicador.day}/{dia_indicador.month}'

#texto do email
mail.HTMLBody = f''' 
<p>Prezados, Bom Dia </p>
<p>A melhor loja do Dia em Faturamento foi a loja {faturamento_lojas_dia.index[0]} com faturamento de R$ {faturamento_lojas_dia.iloc[0,0]:.2f}</p>
<p>A pior loja do Dia em Faturamento foi a loja {faturamento_lojas_dia.index[-1]} com faturamento de R$ {faturamento_lojas_dia.iloc[-1,0]:.2f}</p>
<br>
<p>A melhor loja do Ano em Faturamento foi a loja {faturamento_lojas_ano.index[0]} com faturamento de R$ {faturamento_lojas_ano.iloc[0,0]:.2f}</p>
<p>A pior loja do Ano em Faturamento foi a loja {faturamento_lojas_ano.index[-1]} com faturamento de R$ {faturamento_lojas_ano.iloc[-1,0]:.2f}</p>
<br>
<p> Segue em anexo os Rankings do Ano e do Dia de todas as lojas.</p>

<p>Qualquer dúvida estou a disposição </p>
<p>Att., Jader</p>

'''
#buscando documento dentro da pasta do PC e definindo padrao para enviar por email
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Anual.xlsx'
mail.Attachments.Add(str(attachment)) 
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'
mail.Attachments.Add(str(attachment)) #transformando o caminho em texto para o python entender onde procurar


mail.Send() #email enviado
print(f"O e-mail foi enviado")