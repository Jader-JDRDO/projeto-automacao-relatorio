
import pandas as pd
import win32com.client as win32
import pathlib
from sqlalchemy import create_engine
import sqlite3
import matplotlib.pyplot as plt
import time

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
        if loja not in lista_arquivos_backup:
            nova_pasta = caminho_backup / loja
            nova_pasta.mkdir()
    
        nome_arquivo='{}_{}_{}.xlsx'.format(dia_indicador.month,dia_indicador.day,loja)
        local_arquivo = caminho_backup / loja / nome_arquivo
                  
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
    time.sleep(5)

faturamento_lojas = vendas_df.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
print(faturamento_lojas_ano)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


vendas_dia = vendas_df.loc[vendas_df['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
print(faturamento_lojas_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


engine = create_engine('sqlite:///banco_logistica.db')
try:
      vendas_df.to_sql('tb_vendas_consolidado', if_exists='append', index=False, con=engine)
     
      print("Dados integrados ao SQL com sucesso!")
      
except Exception as e:
    print(f"Erro ao salvar no SQL: {e}")


query_faturamento = """
SELECT 
    Loja,
    SUM([Valor Final]) as [Faturamento Total]
FROM
    tb_vendas_consolidado
GROUP BY 
    Loja
ORDER BY 
    [Faturamento Total] DESC;
"""

with sqlite3.connect('banco_logistica.db') as conn:
    df_faturamento = pd.read_sql(query_faturamento, conn)

cores = []
total_lojas = len(df_faturamento)

for i in range(total_lojas):
    if i < 3:
        cores.append('forestgreen')
    elif i >= (total_lojas - 3):
        cores.append('darkred')
    else:
        cores.append('royalblue')
plt.figure(figsize=(12, 7))#dimensoes do grafico que quero que tenha os dados coletados do top do dia

# Usando gráfico de barras horizontais (barh) para facilitar a leitura dos nomes
barras_a = plt.barh(df_faturamento['Loja'], df_faturamento['Faturamento Total'], color=cores) #variavel barras recebendo os dados do bairro e quantidade de entregas para o eixo x e y do grafico
plt.xticks([]) 


plt.xlim(0, df_faturamento['Faturamento Total'].max() * 1.20)
for barra in barras_a: #para cada barra no dicionario barras
  tamanho = barra.get_width() #pega o valor da quantidade
  plt.text(tamanho + 10000,  #posição x um pouco depois da barra
            barra.get_y() + barra.get_height()/2, #posição y no meio da barra
            f'R$ {int(tamanho):.2f}', #o texto da quantidade convertida em inteiro)
            va='center', fontsize=10, fontweight='bold') #formatacao do texto em centralizado,tamanho 10, negrito
plt.gca().invert_yaxis()
plt.title('Ranking Anual de Faturamento por Loja', fontsize=14) #titulo do grafico com tamanho 14
plt.xlabel('Faturamento Acumulado (R$)') #parte inferior do grafico com a legenda numero de entregas
plt.ylabel('Loja')#parte lateral do grafico com legenda do bairro
plt.tight_layout() #garantindo que todos as informaçoes apareçam no grafico sem areas cortadas ou incompletas
plt.savefig(caminho_backup / 'faturamento_bairro_ano.png')#criando uma figura a partir do texto  
#plt.show()#exibindo a figura formada



    
#enviando email
outlook = win32.Dispatch('outlook.application')
nome_diretoria = emails_df.loc[emails_df['Loja'] ==  'Diretoria', 'E-mail'].values[0] 
mail = outlook.CreateItem(0)
mail.To = 'mrcartonrex@gmail.com'
mail.Subject = f'Ranking {dia_indicador.day}/{dia_indicador.month}'

#texto do email
mail.HTMLBody = f"""
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
"""
#buscando documento dentro da pasta do PC e definindo padrao para enviar por email
attachment_ano  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment_ano))
attachment_dia  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment_dia))
attachment_imagem = pathlib.Path.cwd() / caminho_backup / 'faturamento_bairro_ano.png'
mail.Attachments.Add(str(attachment_imagem))

mail.Send() #email enviado
print(f"O e-mail para a {nome_diretoria} foi enviado")