import cv2
import numpy as np 
import pyqrcode
import PIL
import os
from pyzbar.pyzbar import decode
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import magenta, red
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, Spacer, PageBreak, Table, TableStyle# add this at the top of the script
from reportlab.lib import utils
import pandas as pd
import datetime as dt
import sys
from tkinter import Tk, StringVar, Label, Entry, Button
from tkinter import Canvas, Frame, BOTH, NW
from functools import partial
from PIL import Image, ImageTk
from tkinter import * 

#vARIABLE GLOBALE PERMETTANT D'ETEINDRE LA WEBCAM APRES LA FIN DU SCAN
os.environ["OPENCV_VIDEOIO_PRIORITY_MSMF"] = "0"

################################################################################
################################################################################
###########LA FONCTION QUI PERMET DE SCANNER LE CODE QR VIA LA WEBCAM ##########
def scanner():
	# Capture the video from default camera
	capture = cv2.VideoCapture(0)
	a = 0
	recieved_data = None
	while True:
		# reading frame from the camera
		_, frame = capture.read()
		decoded_data = decode(frame)
		for QR in decode(frame):
			'''pts = np.array([QR.polygon],np.int32)
			pts = pts.reshape((-1,1,2))
			cv2.polylines(frame, [pts],True,(0,255,0),5)'''
			data = decoded_data[0][0]
			data = data.decode()
			if data != recieved_data:
				recieved_data = data
				entrée_data.delete(0, END)
				entrée_data.insert(0,data)
				AnneeUniCertificat()
				capture.release()
				cv2.destroyAllWindows()
				recupere()
				break
		cv2.imshow("QR CODE Scanner", frame)
		# To exit press Esc Key.
		key = cv2.waitKey(1)
		if key == 27:
			capture.release()
			cv2.destroyAllWindows()
			break
################################################################################
################################################################################
################################################################################

################################################################################
################################################################################
##########LA FONCTION QUI PERMET D'IMPRIMER LE CERTIFICAT DE SCOLARITE##########
def imprimer(pdf):
	os.startfile(pdf,"print")
################################################################################
################################################################################
################################################################################

################################################################################
################################################################################
#######LA FONCTION QUI PERMET DE DEDUIRE L'ANNEE UNIVERSITAIRE EN COURS ########
def AnneeUniCertificat():
	entrée_année.delete(0, END)
	date = dt.datetime.now()
	if date.month < 9:
		entrée_année.insert(0,str(date.year-1))
	else:
		entrée_année.insert(0,str(date.year))
################################################################################
################################################################################
################################################################################


fenetre = Tk()


label = Label(fenetre, text="         Université des Sciences et Technologie Houari Boumediene          ", font =('Constancia', 16), bg="Red")
label.pack()
label = Label(fenetre, text="                  Production de Certificat de Scolarité PG                 ", font =('Constancia', 16), bg="Green")
label.pack()
label = Label(fenetre, text="                          Copyright Faculté d'Electronique et Informatique", font =('Constancia', 8), bg="White")
label.pack()
label = Label(fenetre, text="                                   Auteur: Prof. Slimane Larabi, Oct. 2020", font =('Constancia', 8), bg="White")
label.pack()


# bouton de sortie
bouton=Button(fenetre, text="Fermer", command=fenetre.quit)
bouton.pack()

def recupere():
    query=entrée_data.get()
    certificat(entrée_data.get())

value = IntVar() 
value.set("Saisissez l'année universitaire en cours: 2019 pour 2019/2020, 2020 pour 2020/2021")
entrée_année = Entry(fenetre, textvariable=value, width=80, bg="yellow")
entrée_année.pack()

value = IntVar() 
value.set("Valeur")
entrée_année = Entry(fenetre, width=30)
entrée_année.pack()
# entrée
value = StringVar() 
value.set("Saisissez le Nom ou le Matricule du Doctorant")
entrée_data = Entry(fenetre, textvariable=value, width=40, bg="yellow")
entrée_data.pack() 

value = StringVar() 
value.set("Valeur")
entrée_data = Entry(fenetre, width=30)
entrée_data.pack()

#Frame utiliser pour aligner les boutons Scanner et Valider
top = Frame(fenetre)
bottom = Frame(fenetre)
top.pack(side=TOP)
bottom.pack(side=BOTTOM, fill=BOTH, expand=True)

#Bouton utiliser pour scanner le code QR
boutonScan=Button(fenetre, text="Scanner", command=scanner)
boutonScan.pack(in_=top, side=LEFT)

bouton = Button(fenetre, text="Valider", command=recupere)
bouton.pack(in_=top, side=LEFT)

#une Entry non visible par l'utilisateur faite pour conserver le nom du pdf generer utiliser lors de l'impression
pdfname = Entry(fenetre, width=50)
pdfname.pack()
pdfname.pack_forget()
##################################################################################################################

# Create and emplty Label to put the result in
resultLabel = Text(fenetre)



#c = canvas.Canvas("certificat_doctorant.pdf")
numrow=0
data=[[],[]]
img = utils.ImageReader("entete.jpg")
iw, ih = img.getSize()
query=entrée_data.get()


now = dt.datetime.now()
print (now.year)

#Bouton uiliser pour imprimer le pdf
boutonPrint=Button(fenetre, text="Imprimer", command=(lambda: imprimer(pdfname.get())))
#variable utiliser pour creer une exception
affichqr = 0

def certificat(query):
	found=0
	resultLabel.delete(1.0, END)
	numrow=0
	df1 = pd.read_excel ('data.xlsx', sheet_name='Identification')
	for index, row in df1.iterrows():
		nom=row['NOM']
		if(nom.lower() == query.lower()):
			found=1
			numrow=index
			mat=row['Matricule']
			prenom=row['PRENOM']
			sexe=row['Sexe']
			data[0]=[nom,prenom]
			date_naiss=row['DATE DE NAISSANCE']
			lieu_naiss=row['LIEU DE NAISSANCE']
			email=row['Email']
			tel=row['Telephone']
			nat=row['Nationalité']
			resultLabel.insert(END, "Matricule: ")
			resultLabel.pack()
			resultLabel.insert(END, str(mat)+ '\n')
			resultLabel.insert(END, str(nom)+" ")
			resultLabel.insert(END, str(prenom)+ '\n')
			if sexe=="M" :
				resultLabel.insert(END, "Né le:")
			else :
				resultLabel.insert(END, "Née le:")
			resultLabel.insert(END, str(date_naiss.strftime('%d/%m/%Y')))
			resultLabel.insert(END, " à ")
			resultLabel.insert(END, str(lieu_naiss)+ '\n')
			resultLabel.pack()
			resultLabel.insert(END, "email: ")
			resultLabel.insert(END, str(email)+ '\n')
			resultLabel.insert(END, "téléphone : ")
			resultLabel.insert(END, str(tel)+ '\n')
			#resultLabel.insert()
			resultLabel.pack()
			file="certificat_doctorant"+nom+".pdf"
			pdfname.delete(0,END)
			pdfname.insert(0,file)
			c = canvas.Canvas(file)
			print("creation by name", "numrow=", numrow)
	
	if (found == 0) :
		print("name not found")
		for index, row in df1.iterrows():
			mat=row['Matricule']
			if(str(mat).lower() == str(query).lower()):
				found=1
				numrow=index
				nom=row['NOM']
				prenom=row['PRENOM']
				sexe=row['Sexe']
				data[0]=[nom,prenom]
				date_naiss=row['DATE DE NAISSANCE']
				lieu_naiss=row['LIEU DE NAISSANCE']
				email=row['Email']
				tel=row['Telephone']
				nat=row['Nationalité']
				resultLabel.insert(END, "Matricule: ")
				resultLabel.pack()
				resultLabel.insert(END, str(mat)+ '\n')
				resultLabel.insert(END, str(nom)+" ")
				resultLabel.insert(END, str(prenom)+ '\n')
				if sexe=="M" :
					resultLabel.insert(END, "Né le:")
				else :
					resultLabel.insert(END, "Née le:")
				resultLabel.insert(END, str(date_naiss.strftime('%d/%m/%Y')))
				resultLabel.insert(END, " à ")
				resultLabel.insert(END, str(lieu_naiss)+ '\n')
				resultLabel.pack()
				resultLabel.insert(END, "email: ")
				resultLabel.insert(END, str(email)+ '\n')
				resultLabel.insert(END, "téléphone : ")
				resultLabel.insert(END, str(tel)+ '\n')
				resultLabel.pack()
				file="certificat_doctorant"+nom+".pdf"
				pdfname.delete(0,END)
				pdfname.insert(0,file)
				c = canvas.Canvas(file)
				print("creation by Matricule")
				print("matricule found")
				break
	
		
	if( found == 1) :
			
		df2 = pd.read_excel ('data.xlsx', sheet_name='Sujet_Doctorat')
		print("FOUND1")
		for index, row in df2.iterrows():
			if(index==numrow):
				Type_doct=row['Type de Doctorat']
				Filière=row['Filière']
				print(Filière)
				Domaine=row['Domaine']
				Spécialité=row['Spécialité']
				Intitulé=row['Intitule du sujet']
				resultLabel.insert(END, "Doctorat : ")
				resultLabel.insert(END, str(Type_doct)+"\n")
				resultLabel.insert(END, "Domaine : ")
				resultLabel.insert(END, str(Domaine)+"\n")
				resultLabel.insert(END, "Filière : ")
				resultLabel.insert(END, str(Filière)+"\n")
				resultLabel.insert(END, "Spécialité : ")
				resultLabel.insert(END, str(Spécialité)+"\n")
				resultLabel.insert(END, "Intitulé du sujet : ")
				resultLabel.insert(END, str(Intitulé)+"\n")
    
		df3 = pd.read_excel ('data.xlsx', sheet_name='Historique_Inscriptions')
		for index, row in df3.iterrows():
			if(index==numrow):
				année_inscription=row['Année de première  inscription']
				nombre_Gel=row['Gel'] 
				resultLabel.insert(END, "Année de première  inscription: ")
				resultLabel.insert(END, str(année_inscription)+"\n") 
				année=int(entrée_année.get())-int(année_inscription)+1-int(nombre_Gel)
				print(année)
				année_inscription=année

		df4 = pd.read_excel ('data.xlsx', sheet_name='Directeur_thèse')
		print("FOUND3")
		for index, row in df4.iterrows():
			if(index==numrow):
				Directeur=row['Nom et Prénom du Directeur de thèse'] 
				Grade=row['Grade du Directeur de thèse']
				resultLabel.insert(END, "Directeur de thèse : ")
				resultLabel.insert(END, str(Directeur)+",")
				resultLabel.insert(END, str(Grade)+"\n")
				coDirecteur=row['Nom et Prénom du co-Directeur de thèse'] 
				coGrade=row['Grade du co-Directeur de thèse']
				resultLabel.insert(END, "co-Directeur de thèse : ")
				resultLabel.insert(END, str(coDirecteur)+",")
				resultLabel.insert(END, str(coGrade)+"\n")    

##################################################################################
##################################################################################
##lA PARTIE DU CODE QUI S'OCCUPE D'AFFICHER LE CODE QR SUR L'INTERFACE GRAPHIQUE## 
		qr = pyqrcode.create(mat) 
		nomQR = mat+'.png'
		qr.png(nomQR,scale = 5)
		img = cv2.imread(nomQR)
		qrimg = cv2.resize(img,(80,80))
		img1 = ImageTk.PhotoImage(image=PIL.Image.fromarray(qrimg))
		resultLabel.insert(END, "                                                                     ")
		resultLabel.image_create(END, image = img1)
##################################################################################
##################################################################################

		c.drawImage("entete.jpg", 0, 10*inch, iw*0.6, ih*0.6, preserveAspectRatio=True)

##################################################################################
##################################################################################
##########lA PARTIE DU CODE QUI S'OCCUPE D'AFFICHER LE CODE QR SUR LE PDF#########
		qri = utils.ImageReader(nomQR)
		qw, qh = qri.getSize()
		c.drawImage(nomQR, 50, 100, 80, 80, preserveAspectRatio=True)
##################################################################################
##################################################################################

		c.setFont("Helvetica-Bold", 16)
		c.drawCentredString(4*inch, 9.8*inch, "CERTIFICAT DE SCOLARITE")
		c.drawCentredString(4*inch, 9.4*inch, "POST-GRADUATION")
		c.setFont("Helvetica", 11)
		c.drawString(inch, 8.5*inch, "Le Doyen de la Faculté d'Electronique et d'Informatique de l'Université des Sciences")
		c.drawString(inch, 8.2*inch, "et de la Technologie Houari Boumediene Certifie que :")
		#c.drawString(inch, 7.9*inch, "")
		c.setFont("Helvetica-Bold", 11)

		if sexe=="M" :
			c.drawString(1.2*inch, 7.5*inch, "Mr:")
		else :
			c.drawString(1.2*inch, 7.5*inch, "Melle :")
		c.setFont("Helvetica", 11)
	#c.drawString(2.0*inch, 7.5*inch, nom)
	#c.drawString(2.6*inch, 7.5*inch, prenom)
		t=Table(data, colWidths=None)
		t.setStyle(TableStyle([('ALIGN', (0,0), (1, 1), 'LEFT')]))
		aW = 460 # available width and height
		aH = 900
		w,h = t.wrap(aW, aH)
		w=2*inch
		h=7.18*inch
  
		c.setFont("Helvetica-Bold", 11)
		t.drawOn(c, w, h)#,2*inch, 5*inch)
		if sexe=="M" :
			c.drawString(1.2*inch, 7.1*inch, "Né le :")
		else :
			c.drawString(1.2*inch, 7.1*inch, "Née le :")
		c.setFont("Helvetica", 11)
		c.drawString(2.0*inch, 7.1*inch, date_naiss.strftime('%d/%m/%Y'))
		c.drawString(3.2*inch, 7.1*inch, "à :")
		c.drawString(3.5*inch, 7.1*inch, lieu_naiss)
		c.setFont("Helvetica-Bold", 11)
		c.drawString(1.2*inch, 6.7*inch, "Nationalité :")
		c.setFont("Helvetica", 11)
		c.drawString(2.2*inch, 6.7*inch, nat)
		c.setFont("Helvetica-Bold", 11)
		c.drawString(1.2*inch, 6.3*inch, "Matricule :")
		c.setFont("Helvetica", 11)
		c.drawString(2.2*inch, 6.3*inch, mat)
			
		if Type_doct=="LMD" :
			c.setFont("Helvetica-Bold", 11)
			c.drawString(1.2*inch, 5.9*inch, "Domaine :")
			if(Domaine=="MI"):
				c.setFont("Helvetica", 11)
				c.drawString(2.2*inch, 5.9*inch, "Mathématiques Informatique")
			else:
				c.setFont("Helvetica", 11)
				c.drawString(2.2*inch, 5.9*inch, "Sciences et Technologies")
			c.setFont("Helvetica-Bold", 11)
			c.drawString(1.2*inch, 5.5*inch, "Filière :")
			c.setFont("Helvetica", 11)
			c.drawString(2.2*inch, 5.5*inch, Filière)
			c.setFont("Helvetica-Bold", 11)
			c.drawString(1.2*inch, 5.1*inch, "Spécialité :")
			c.setFont("Helvetica", 11)
			c.drawString(2.2*inch, 5.1*inch, Spécialité)
			if sexe=="M" :
				c.drawString(1.0*inch, 4.7*inch, "est inscrit  au titre de l'année universitaire ")
			else :
				c.drawString(1.0*inch, 4.7*inch, "est inscrite  au titre de l'année universitaire ")
			c.drawString(4.0*inch, 4.7*inch, str(int(entrée_année.get())))
			c.drawString(4.35*inch, 4.7*inch, "/")
			c.drawString(4.4*inch, 4.7*inch, str(int(entrée_année.get())+1))

			c.drawString(1.0*inch, 4.3*inch, "En: ")
			c.drawString(1.3*inch, 4.3*inch, str(année_inscription))
			if année_inscription <10 :
				c.drawString(1.4*inch, 4.3*inch, "ème Année")
			else :
				c.drawString(1.5*inch, 4.3*inch, "ème Année")
			c.drawString(2.35*inch, 4.3*inch, "Doctorat.")
						
		else:
			c.setFont("Helvetica-Bold", 11)
			c.drawString(1.2*inch, 5.9*inch, "Filière :")
			c.setFont("Helvetica-Bold", 11)
			c.drawString(2.2*inch, 5.9*inch, Filière)
			c.setFont("Helvetica", 11)
			c.drawString(1.2*inch, 5.5*inch, "Spécialité :")
			c.setFont("Helvetica-Bold", 11)
			c.drawString(2.2*inch, 5.5*inch, Spécialité)
			if sexe=="M" :
				c.drawString(1.0*inch, 5.1*inch, "est inscrit  au titre de l'année universitaire ")
			else :
				c.drawString(1.0*inch, 5.1*inch, "est inscrite  au titre de l'année universitaire ")
			c.drawString(4.0*inch, 5.1*inch, str(int(entrée_année.get())))
			c.drawString(4.35*inch, 5.1*inch, "/")
			c.drawString(4.4*inch, 5.1*inch, str(int(entrée_année.get())+1))

			c.drawString(1.0*inch, 4.7*inch, "En: ")
			c.drawString(1.3*inch, 4.7*inch, str(année_inscription))
			if année_inscription <10 :
				c.drawString(1.4*inch, 4.7*inch, "ème Année")
			else :
				c.drawString(1.5*inch, 4.7*inch, "ème Année")
			c.drawString(2.35*inch, 4.7*inch, "Doctorat en Sciences.")
			
	
		c.drawString(4.7*inch, 3.3*inch, "Bab-Ezzouar, le :")
		c.drawString(6.0*inch, 3.3*inch, dt.datetime.today().strftime("%d/%m/%Y"))
		c.setFont("Helvetica", 11)
		c.drawString(4.9*inch, 3.0*inch, "Le Doyen")
		c.setFont("Helvetica-Bold", 10)
		c.drawCentredString(4*inch, 1.0*inch,"Faculté d'Electronique et d'Informatique")
		c.setFont("Helvetica", 10)
		c.drawCentredString(4*inch, 0.8*inch,"USTHB, BP. 32, El Alia, Bab Ezzouar 16111, Alger")
		c.drawCentredString(4*inch, 0.6*inch,"Tel: +213023934066, Fax: +213023934066, email: pgfei@usthb.dz") 
		c.setFont("Helvetica", 8)
		c.line(0, 1.2*inch, 10*inch, 1.1*inch)
		c.showPage()
		if ((année_inscription <= 6) and (Type_doct !="LMD")) or ((année_inscription <= 5) and (Type_doct =="LMD")):
			c.save()
		else:
			print("Doctorant n'ayant pas droit au certificat de scolarité")
			resultLabel.insert(END, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n")
			resultLabel.pack()
			resultLabel.insert(END, "                              \n")
			resultLabel.pack()
			resultLabel.insert(END, " Doctorant n'ayant pas droit  \n")
			resultLabel.pack()
			resultLabel.insert(END, " au certificat de scolarité   \n")
			resultLabel.pack()
			resultLabel.insert(END, "                              \n")
			resultLabel.pack()
			resultLabel.insert(END, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n")
			resultLabel.pack()
			
		#Cette condition est la afin d'eviter qu'un nouveau bouton soit generer a chaque creation d'un nouveau pdf
		if boutonPrint.winfo_ismapped() == False:
			boutonPrint.pack()
		
		#Pour certaines raison le code QR s'affiche sur l'interface graphique uniquement lorsqu'une exception a lieu
		# dans la methode certifie c'est pour cela que j'ai mis cette condition generant une exception qui ne pose pas
		# de probleme lors de l'execution du programme 
		if affichqr == 0:
			affichqr = 1



	else:
		print("matriculeand name not found")
		resultLabel.insert(END, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n")
		resultLabel.pack()
		resultLabel.insert(END, "                              \n")
		resultLabel.pack()
		resultLabel.insert(END, " Doctorant n'existe pas dans  \n")
		resultLabel.pack()
		resultLabel.insert(END, "        la base data.xlsx     \n")
		resultLabel.pack()
		resultLabel.insert(END, "                              \n")
		resultLabel.pack()
		resultLabel.insert(END, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n")
		resultLabel.pack()

fenetre.mainloop()