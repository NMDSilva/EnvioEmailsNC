# pip install pywin32
import win32com.client as win32
# pip install dbfread
from dbfread import DBF
# table = DBF('lista.dbf', load=True)
import os
from pathlib import Path
# pip install win10toast
from win10toast import ToastNotifier
toaster = ToastNotifier()

outlook = win32.Dispatch('outlook.application')

caminho = str(Path().absolute())
ficheirosEnviados = Path('enviados')
ficheirosEnviados.mkdir(exist_ok=True)

fileTxt = open('assunto.txt', mode='r', encoding='utf-8')
assunto = fileTxt.read()
fileTxt.close()

fileHtml = open('email.html', mode='r', encoding='utf-8')
emailTxt = fileHtml.read()
fileHtml.close()

contador = 0

def getDados(nif, campo):
  table = DBF('lista.dbf', load=True)
  for x in table.records:
    if str(nif) == str(x['nif']):
      return x[str(campo)]
  return ''

def enviarEmail(nome, email, ficheiro):
  try:
    mail = outlook.CreateItem(0)
    assinatura = mail.Attachments.Add(caminho + '\\logo.png')
    assinatura.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', 'assinatura')
    mail.Attachments.Add(caminho + '\\' + str(ficheiro))
    mail.To = str(email)
    mail.Subject = str(assunto)
    mail.HTMLBody = emailTxt.replace('{{nome}}', nome).replace('{{email}}', email)
    mail.Send()
    os.rename(caminho + '\\' + ficheiro, str(ficheirosEnviados.absolute()) + '\\' + ficheiro)
    global contador
    contador += 1
  except:
    toaster.show_toast('Problema ao enviar fatura', 'Problemas ao enviar o ficheiro ' + str(ficheiro))

for ficheiro in os.listdir('.'):
  if ficheiro.endswith('.pdf'):
    nif = str(ficheiro.replace('.pdf', ''))
    nome = getDados(nif, 'nome')
    if nome == '':
      toaster.show_toast('Cliente não existe', 'O NIF ' + nif + ' não existe na base de dados')
      continue
    email = getDados(nif, 'email')
    if email == '':
      toaster.show_toast('Cliente sem email registado', 'O cliente com o NIF ' + nif + ' não tem email registado')
      continue
    enviarEmail(nome, email, ficheiro)

toaster.show_toast('Envio finalizado', 'Programa terminou o envio das faturas. Faturas enviadas: ' + str(contador))
