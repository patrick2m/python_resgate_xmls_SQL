import pyodbc
import os
from datetime import datetime
from configuracoes import SERVER_IP, SERVER_DATABASE, SERVER_NAME, SERVER_PASSWORD

########### Busca informações do servidor para conectar à base de dados ###########

server = SERVER_IP
database = SERVER_DATABASE
username = SERVER_NAME
password = SERVER_PASSWORD
cnxn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER='+server+';DATABASE='+database+';ENCRYPT=yes;UID='+username+';PWD='+ password+';TrustServerCertificate=Yes;')
cursor = cnxn.cursor()

########### PARAMETROS ###########

# Data pra puxar desde :

datainicial = '20230101'
datafinal = '20230131'

# Caminho da pasta inicial + nome da pasta criada #

diretorio = 'c:\\AAAAAAAAA\\' + datainicial + ' até ' + datafinal

########### EXPORTAR XMLS DAS SÉRIES DA CASA DAS FECHADURAS DE NITERÓI ##########

seriesCFN = ['CF1', '101', '102']

for i in range(0, len(seriesCFN)):
  os.makedirs(os.path.join(diretorio + '\\CFN', str(seriesCFN[i])))

def criarArquivosXml(serie, datainicial, datafinal):
  cursor.execute("SELECT * FROM TSS.dbo.SPED050 WHERE D_E_L_E_T_ = '' AND SUBSTRING(NFE_ID,1,3) = ? AND DATE_NFE >= ? AND DATE_NFE <= ? ", (serie,datainicial, datafinal))
  row = cursor.fetchall()
  for i in range(0, len(row)):
    if len(str(row[i][5])[2:-5]) > 0:
      with open(diretorio + '\\CFN\\' + serie + '\\' + str(row[i][1])[0:12] + '.xml', 'a') as j:
        j.write(str(row[i][5])[2:-5])
  row = ''
for i in range(0, len(seriesCFN)):
  criarArquivosXml(seriesCFN[i],datainicial, datafinal)

########### EXPORTAR XMLS DAS SÉRIES DA CASA DAS FECHADURAS DE SÃO GONÇALO ##########

seriesCFSG = ['SG1', '001']

for i in range(0, len(seriesCFSG)):
  os.makedirs(os.path.join(diretorio + '\\CFSG', str(seriesCFSG[i])))

def criarArquivosXml(serie, datainicial, datafinal):
  cursor.execute("SELECT * FROM TSS.dbo.SPED050 WHERE D_E_L_E_T_ = '' AND SUBSTRING(NFE_ID,1,3) = ? AND DATE_NFE >= ? AND DATE_NFE <= ? ", (serie,datainicial, datafinal))
  row = cursor.fetchall()
  for i in range(0, len(row)):
    if len(str(row[i][5])[2:-5]) > 0:
      with open(diretorio + '\\CFSG\\' + serie + '\\' + str(row[i][1])[0:12] + '.xml', 'a') as j:
        j.write(str(row[i][5])[2:-5])
  row = ''
for i in range(0, len(seriesCFSG)):
  criarArquivosXml(seriesCFSG[i],datainicial, datafinal)

########### EXPORTAR XMLS DAS SÉRIES DA BELLO BANHO ##########

seriesBB = ['BB1', '090']

for i in range(0, len(seriesBB)):
  os.makedirs(os.path.join(diretorio + '\\BB', str(seriesBB[i])))

def criarArquivosXml(serie, datainicial, datafinal):
  cursor.execute("SELECT * FROM TSS.dbo.SPED050 WHERE D_E_L_E_T_ = '' AND SUBSTRING(NFE_ID,1,3) = ? AND DATE_NFE >= ? AND DATE_NFE <= ? ", (serie,datainicial, datafinal))
  row = cursor.fetchall()
  for i in range(0, len(row)):
    if len(str(row[i][5])[2:-5]) > 0:
      with open(diretorio + '\\BB\\' + serie + '\\' + str(row[i][1])[0:12] + '.xml', 'a') as j:
        j.write(str(row[i][5])[2:-5])
  row = ''
for i in range(0, len(seriesBB)):
  criarArquivosXml(seriesBB[i],datainicial, datafinal)

########### EXPORTAR XMLS DAS SÉRIES DA ISAAC DAS FERRAMENTAS ##########

seriesISAAC = ['IS1', '201']

for i in range(0, len(seriesISAAC)):
  os.makedirs(os.path.join(diretorio + '\\ISAAC', str(seriesISAAC[i])))

def criarArquivosXml(serie, datainicial, datafinal):
  cursor.execute("SELECT * FROM TSS.dbo.SPED050 WHERE D_E_L_E_T_ = '' AND SUBSTRING(NFE_ID,1,3) = ? AND DATE_NFE >= ? AND DATE_NFE <= ? ", (serie,datainicial,datafinal))
  row = cursor.fetchall()
  for i in range(0, len(row)):
    if len(str(row[i][5])[2:-5]) > 0:
      with open(diretorio + '\\ISAAC\\' + serie + '\\' + str(row[i][1])[0:12] + '.xml', 'a') as j:
        j.write(str(row[i][5])[2:-5])
  row = ''
for i in range(0, len(seriesISAAC)):
  criarArquivosXml(seriesISAAC[i],datainicial,datafinal)