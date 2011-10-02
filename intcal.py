# -*- coding: utf-8 -*-

import re
import sys
import os
import xlrd
import icalendar
from icalendar import Calendar, Event
from datetime import datetime

#Version du programme
version = 0.3

#Priorité des événements
priority = 3

class Intcal:
	def chargerFichier(self, fichier):
		"""Charge le fichier xls dans un objet icalendar"""
		try:
			self.book = xlrd.open_workbook(fichier)
		except:
			print "Erreur, le fichier xls est incorrect"
	
	def extraire_titre(self):
		"""Extrait le titre du cours"""
		if (len(self.texte_nl) >= 1):
			self.titre = self.texte_nl[0]
			return True
		else:
			self.titre = ""
			return False

	def extraire_groupe(self):
		"""Extrait les groupes du cours"""
		self.groupes = []
		for i in range(1, len(self.texte_nl)):
			regex = re.search("Gp-", self.texte_nl[i])
			if (regex):
				self.groupes.append(self.texte_nl[i])

	def extraire_type(self):
		"""Extrait les types (Cours Magistral, TP, TD, ...) du cours"""
		self.types = []
		for i in range(1, len(self.texte_nl)):
			regex = re.search("TP|TD|Examen|Conférence|Cours Magistral|Présentation|Point de Rencontre|Projet|Formation en ligne|Soutenance", self.texte_nl[i].encode("utf-8"))
			if (regex and not self.texte_nl[i] in self.groupes):
				self.types.append(self.texte_nl[i])
			else:
				#Les types sont affichés à la suite. Si la dernière ligne n'en n'est pas un, on arrête.
				break
	
	def extraire_date(self):
		"""Extrait la date et les horaires du cours"""
		self.date = False
		self.horaire_deb = False
		self.horaire_fin = False
		for i in range(1, len(self.texte_nl)):
			regex = re.search("^[0-9]{2}/[0-9]{2}/[0-9]{4}$", self.texte_nl[i])
			if (regex):
				self.date = regex.group(0)
			regex = re.search("^([0-9]{2})h([0-9]{2})-([0-9]{2})h([0-9]{2})$", self.texte_nl[i])
			if (regex):
				self.horaire_deb = [regex.group(1), regex.group(2)]
				self.horaire_fin = [regex.group(3), regex.group(4)]
			if (self.date and self.horaire_deb and self.horaire_fin):
				break
		if (not self.date or not self.horaire_deb or not self.horaire_fin):
			return False
		else:
			return True

	def extraire_iuff(self):
		"""Extrait l'IUFF/CUFF/IUAF du cours"""
		self.iuff = False
		for i in range(1, len(self.texte_nl)):
			regex = re.search("^(IUFF|CUFF|IUAF).*", self.texte_nl[i])
			if (regex):
				self.iuff = regex.group(0)
				break

	def extraire_description(self):
		"""Extrait la description du cours"""
		self.description = []
		i = len(self.texte_nl)-1
		while (not self.texte_nl[i] in self.groupes):
			self.description.append(self.texte_nl[i])
			i-=1

	def extraire_salle(self):
		"""Extrait les salles du cours"""
		self.salles = []
		for i in range(0, len(self.texte_nl)):
			regex = re.search("(A|B|C|D|E|F|BL)[0-9]+.*|(SSP|FORUM|GYMNASE|AMPHI).*", self.texte_nl[i])
			if (regex):
				self.salles.append(regex.group(0))
	
	def extraire_formateur(self):
		"""Extrait les formateurs du cours"""
		self.formateurs = []
		preDepart = False
		depart = False
		for i in range(1, len(self.texte_nl)):
			if ((self.texte_nl[i] in self.salles) or (self.texte_nl[i] in self.groupes)):
				#On est trop loin
				break
			elif (depart):
				self.formateurs.append(self.texte_nl[i])
			if (self.texte_nl[i] == self.iuff or preDepart):
				depart = True
			if (self.texte_nl[i] == "ACT"):
				preDepart = True

	def cal_entete(self):
		self.cal = Calendar()
		self.cal.add('prodid', '-//INTCal//')
		self.cal.add('version', str(version))
	
	def cal_creer_evenement(self):
		self.event = Event()
		if (self.titre != "" and self.date and self.horaire_deb and self.horaire_fin):
			self.event.add('summary', self.titre)
			date_split = self.date.split("/")
			self.event.add('dtstart', datetime(int(date_split[2]), int(date_split[1]), int(date_split[0]), int(self.horaire_deb[0]), int(self.horaire_deb[1]), 0))
			self.event.add('dtend', datetime(int(date_split[2]), int(date_split[1]), int(date_split[0]), int(self.horaire_fin[0]), int(self.horaire_fin[1]), 0))
			self.event.add('dtstamp',  datetime.now())
			self.event.add('priority', priority)
			self.event.add('location', " ".join(self.salles))
			self.event.add('description', "\N".join(self.types)+"\N"+self.iuff+"\N"+"\N".join(self.formateurs)+"\N"+" ".join(self.groupes)+"\N"+"\N".join(self.description))
			self.cal.add_component(self.event)

	def cal_ecrire(self, fichier):
		try:
			f = open(fichier, 'wb')
			f.write(self.cal.as_string())
			f.close()
		except:
			print "Impossible d'écrire dans le fichier ics, veuillez vérifier vos droits"

	def parcourir(self):
		"""Parcourt les cellules du fichier xls"""
		for i in range(self.book.nsheets):
			self.sh = self.book.sheet_by_index(i)
			regex = re.search("[0-9]{2}/[0-9]{2}/[0-9]{4}", self.sh.cell_value(rowx=0, colx=0))
			if (regex):
				self.dateEdition = regex.group(0)
			for rx in range(12, self.sh.nrows):
				for ry in range(1, self.sh.ncols):
					if (self.sh.cell_value(rowx=rx, colx=ry)):
						self.texte = self.sh.cell_value(rowx=rx, colx=ry)
						#print self.texte
						self.texte_nl = self.texte.split("\n")
						if (not self.extraire_titre()):
							print "Titre non trouvé pour \n"+self.texte.encode("utf-8")
						self.extraire_groupe()
						if (len(self.groupes) == 0):
							print "Groupe non trouvé pour \n"+self.texte.encode("utf-8")
						self.extraire_type()
						if (len(self.types) == 0):
							print "Type non trouvé pour \n"+self.texte.encode("utf-8")
						if (not self.extraire_date()):
							print "Horaires non trouvés pour \n"+self.texte.encode("utf-8")
						self.extraire_iuff()
						if (not self.iuff):
							print "IUFF/CUFF/IUAF non trouvés pour \n"+self.texte.encode("utf-8")
						self.extraire_description()
						self.extraire_salle()
						if (len(self.salles) == 0):
							print "Salle non trouvée pour \n"+self.texte.encode("utf-8")
						self.extraire_formateur()
						if (len(self.formateurs) == 0):
							print "Formateur non trouvé pour \n"+self.texte.encode("utf-8")
						self.cal_creer_evenement()

print "INTCal Version "+str(version)
print "ATTENTION : La fraicheur de votre fichier .ics sera celle de votre fichier .xls"
print "Merci de rapporter les bugs à netantho@minet.net en incluant votre emploi du temps au format .xls\n"
if len(sys.argv) <= 2 or not re.search("\.xls", sys.argv[1]) or not re.search("\.ics$", sys.argv[2]):
	print "Utilisation : python ./intcal.py monfichier.xls monfichier.ics"
	print "monfichier.xls est le fichier excel exporté de gaspar tous les détails de votre emploi du temps"
	print "monfichier.ics est le chemin du fichier dans lequel vous voulez enregistrer votre emploi du temps en format ics"
else:
	intcal= Intcal()
	intcal.chargerFichier(sys.argv[1])
	intcal.cal_entete()
	intcal.parcourir()
	intcal.cal_ecrire(sys.argv[2])