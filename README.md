📊 Automação de Indicadores e Ranking de Faturamento por Loja

Este projeto consiste em uma automação completa de Inteligência de Negócios (BI) desenvolvida em Python. O sistema realiza a extração, tratamento, carga em banco de dados SQL (ETL) e distribuição diária de relatórios de performance (OnePages) para os gerentes de 25 lojas, finalizando com um relatório executivo consolidado para a Diretoria.


🚀 Funcionalidades Principais

  ETL Automatizado: Consolida bases de vendas (`.xlsx`), informações de cadastro de lojas (`.csv`) e mapeamento de gerentes (`.xlsx`) utilizando Pandas.
  Backup Inteligente: Separa os dados de vendas de cada loja e gera arquivos específicos em pastas organizadas por estabelecimento.
  Carga em Banco de Dados (SQL):Alimenta um banco de dados SQLite para armazenamento histórico das vendas consolidadas.
  Análise Visual de Performance (Matplotlib): Gera um gráfico de barras horizontais personalizado que destaca instantaneamente o Top 3 (maiores faturamentos em verde) e os 3 mais baixos (menores faturamentos em vermelho).
  Disparos de E-mails em Massa (Outlook): Envia OnePages formatados em HTML com tabelas de metas vs. cenários reais para os 25 gerentes e um e-mail com os rankings e gráficos anexados para a Diretoria.

📊 Visão Geral do Ranking Anual

O gráfico gerado pelo script auxilia o tomador de decisão a identificar rapidamente os extremos de performance da rede de lojas, otimizando o tempo de análise da diretoria:

![Ranking de Faturamento por Loja](Backup%20Arquivos%20Lojas/faturamento_bairro_ano.png)


🛠️ Tecnologias Utilizadas

- Python 3
- Pandas: Manipulação e cruzamento de dados.
- SQLAlchemy & SQLite3: Persistência de dados e consultas analíticas (`GROUP BY`).
- Matplotlib: Construção e estilização de gráficos corporativos.
- PyWin32 (win32com.client): Integração e automação com o Microsoft Outlook.
- Pathlib: Manipulação segura de caminhos e diretórios do sistema operacional.
- Time: Manipulação de tempo para nao sobrecarregar as execuções automaticas e o email não constar como spam


📦 Como Rodar o Projeto

1. Certifique-se de ter o Mozila Firefox instalado e configurado na sua máquina.
2. Certifique-se de estar conectado ao Gmail e ao Google Drive.
3. Clone este repositório com o arquivo python uma pasta de sua preferencia
4. Instale as bibliotecas necessárias:
     pip install pandas pywin32 matplotlib sqlalchemy sqlite3 time
5. Certifique-se que o seu arquivo '.xlsx' ou '.csv' esteja localizado em uma pasta chamada "Exportar" na base onde esta o arquivo pyhton a ser usado
6. Executar normalmente (geralmente 2 cliques)
     (Vale ressaltar que o a resolução a ser usada originalmente é de '1440 x 900', portando, caso aconteça algum erro de interação, ele se deve ao fato de ser uma resolução distinta da          original)

 
