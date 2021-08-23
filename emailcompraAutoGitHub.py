import time
import win32com.client as win32
import os
import string
import tkinter as tk
import re
from tkinter import messagebox
from tkinter import *
from tkinter import ttk
import sys

outlook = win32.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
inbox_enviado = mapi.GetDefaultFolder(5)
messages = inbox_enviado.Items

window = tk.Tk()
window.title("Enviador de E-mail de Compra")


frame1 = tk.Frame(borderwidth=3,relief = "raised", height=10,width=20)
frame1.grid(column=0,row=0, sticky="n")


p1 = tk.Label(frame1,text='OV',wraplength=230,justify='left', font='bold')
p1.grid(column=0,row=0)
p2 = tk.Entry(frame1, width = 10,relief='sunken')
p2.insert(END, 'BR')
p2.grid(column=1,row=0)
p3 = tk.Label(frame1,text='Prazo de entrega',wraplength=230,justify='center',font='bold')
p3.grid(column=0,row=1)
p4 = tk.Entry(frame1, width = 10,relief='sunken')
p4.grid(column=1,row=1)
p5 = tk.Label(frame1,text='Item(ns)',wraplength=230,justify='center',font='bold')
p5.grid(column=0,row=2)
p6 = tk.Entry(frame1, width = 10,relief='sunken')
p6.grid(column=1,row=2)


p7 = tk.Label(frame1,text='Observações:',wraplength=230,justify='center',font='bold')
p7.grid(column=2,row=0)

opções_posições = ['Neste pedido é somente serpentina', 'Frequência da obra em 50Hz', 'Ventilador anti-centelhante e motor a prova de explosão', 'Motor WEG']

entry1 = ttk.Combobox(frame1, values = opções_posições, state='readonly', width = 50, height = 30)
entry1.grid(column=2,row=1)



regexOV = r'BR\d{7}'

var2 = tk.StringVar()
var2.set(0)
Apertou = 0


def BotaoObs():
	global p8
	p8 = tk.Text(frame1, height=3,width=20)
	p8.grid(column=2,row=1,padx=10)
	entry1.destroy()
	global Apertou
	Apertou = 1
	window.geometry("380x145")

r3 = tk.Button(frame1, text="Inserir texto em observações", command = BotaoObs, justify = "left")
r3.grid(column=0,row=3, columnspan = 2, pady=5)

def MandadorEmail():
	#Criação do e-mail
	OV = p2.get()
	PrazoOV = p4.get()
	Item = p6.get()

	if Apertou == 1:
		Observação = p8.get("1.0",'end-1c')
	else:
		Observação = str(entry1.get())
	

	try:
		y = re.search(regexOV,OV).group(0)
	except AttributeError:
		y = re.search(regexOV, OV)
	if y == None:
		tk.messagebox.showinfo("Erro OV", "OV está no formato errado")
		sys.exit()
	else:
		pass

	caminho = 'Confidential'
	caminhoOV = str(caminho + '/' + OV[0:6])

	try:
		subOV = next(os.walk(caminhoOV))[1]
	except StopIteration:
		tk.messagebox.showinfo("OV não encontrada", "Não foi possível encontrar esta OV")
		sys.exit()

	for x in range(len(subOV)):
		if subOV[x].startswith(OV):
			OVcliente = subOV[x]
		else:
			pass

	caminhoOVcliente = str(caminho + '/' + OV[0:6] + '/' + OVcliente + '/DOC')

	mail = outlook.CreateItem(0)
	mail.To ='Confidential'
	mail.Subject = 'COMPRA - REV.0 / OV: ' + OVcliente
	mail.Body = 'teste'

	 
	Parte1 = r'''

	<h1 style="font-size:20px;"><small>
	Bom dia a todos,
	</h1></small>

	<h2>LIBERAÇÃO DE COMPRAS:</h2>

	<li>OV: '''

	Parte2 = r'''</li>
	<li>Cliente: '''

	Parte3 = r'''</li>
	<li>Prazo da OV: '''

	Parte4 = r'''</li>
	<li>Revisão: 0</li>
	<li>Item: '''

	Parte5 = r'''</li>

	<h1 style="font-size:30px;">OBSERVAÇÕES:</h1>


	'''

	cliente = OVcliente.split(' - ')

	mail.HTMLBody = Parte1 + OV + Parte2 + cliente[1] + Parte3 + PrazoOV + Parte4 + Item + Parte5 + Observação

	mail.Send()

	#Fim da criação do e-mail


	#Esperou
	time.sleep(10)
	#Abriu
	Ultimoemail = messages.GetLast()
	Ultimoemailstr = str(messages.GetLast())

	valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)

	UltimoEmailArquivo = ''.join(c for c in Ultimoemailstr if c in valid_chars)

	ArquivoUltimoEmail = str(caminhoOVcliente +'/' + UltimoEmailArquivo + '.msg')

	os.startfile(os.path.realpath(caminhoOVcliente))

	#Salvou
	Ultimoemail.SaveAs(ArquivoUltimoEmail)
	sys.exit()


Mandar_email = tk.Button(frame1,text="Enviar E-mail", command=MandadorEmail,width = 10)
Mandar_email.grid(column=2,row=3)

window.geometry("517x117")
window.eval('tk::PlaceWindow . center')
window.configure(background='black')

tk.mainloop()
