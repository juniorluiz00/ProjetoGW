import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import pyodbc

# Configurações do ODBC
dsn = 
username = 
password = 
sql_query = """
SELECT s.* 
FROM SLDOLOTE s 
JOIN produtos p ON s.cod_item = p.cod_item 
WHERE s.qtd_saldo > '0' 
  AND s.COD_EMPRESA = 'GREEN' 
  AND p.flg_tpitem = 'M'
  AND s.COD_AREAESTOQUE IN ('AGUARDANDO', 'ALMOXARIFADO', 'RESERVA', 'RESERVA_PROCESSO')
"""

# Estabelecer a conexão com o banco de dados
conn = pyodbc.connect('DSN=' + dsn + ';UID=' + username + ';PWD=' + password)

# Executar a consulta SQL e carregar os resultados em um DataFrame
df = pd.read_sql_query(sql_query, conn)

# Fechar a conexão com o banco de dados
conn.close()

# Selecionar apenas as colunas desejadas
colunas_desejadas = ['COD_ITEM', 'COD_AREAESTOQUE', 'NOM_LOTE', 'QTD_SALDO', 'DAT_VALIDADE']
df = df[colunas_desejadas].copy()

# Formatar a data de validade no padrão brasileiro
df['DAT_VALIDADE'] = pd.to_datetime(df['DAT_VALIDADE']).dt.strftime('%d/%m/%Y')

# Ordenar o DataFrame pelo campo 'DAT_VALIDADE' em ordem crescente
df = df.sort_values('DAT_VALIDADE')

# Renomear as colunas
nomes_colunas = ['Código Item', 'Código Área Estoque', 'Nome Lote', 'Quantidade Saldo', 'Data Validade']
df.columns = nomes_colunas

# Caminho do arquivo Excel
caminho_arquivo = r''

# Salvar o DataFrame em um novo arquivo Excel com as colunas e formato desejados
df.to_excel(caminho_arquivo, index=False)

# Configurações de e-mail
remetente = ''
senha = ''
destinatarios = ['']
assunto = 'Materiais com Validade Próxima'
data_atual = datetime.now().strftime('%d/%m/%Y')
data_limite = (datetime.now() + timedelta(days=90)).strftime('%d/%m/%Y')
corpo_email = f"""Olá,

Segue anexo o arquivo com as informações atualizadas sobre os materiais com validade nos próximos três meses.

Data de envio: {data_atual}
Data de vencimento das matérias-primas: {data_atual} a {data_limite}

Atenciosamente,
Seu Nome
"""

# Criação do e-mail
msg = MIMEMultipart()
msg['From'] = remetente
msg['To'] = ', '.join(destinatarios)
msg['Subject'] = assunto

msg.attach(MIMEText(corpo_email, 'plain'))

# Anexando o arquivo ao e-mail
with open(caminho_arquivo, 'rb') as anexo:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(anexo.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {caminho_arquivo}')
    msg.attach(part)

# Conexão com o servidor SMTP e envio do e-mail
try:
    servidor_smtp = ''
    porta_smtp = 

    with smtplib.SMTP_SSL(servidor_smtp, porta_smtp) as servidor:
        servidor.login(remetente, senha)
        servidor.sendmail(remetente, destinatarios, msg.as_string())
    print('E-mail enviado com sucesso!')
except Exception as e:
    print(f'Erro ao enviar e-mail: {str(e)}')