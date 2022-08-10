#install pywin32 en pandas
import win32com.client
import pandas as pd
from string import Template
from datetime import date
import time
import os
#bewaar de mogelijke HTML-teksten in een ander script en importeer ze
from mailberichtteksten import *

#kies de ingevulde csv met Title, firstname, lastname, email, pathname, objectenlijst, toestemming en IERnummer kolommen
#csv moet al ingevuld zijn, zie script "BijlagenOphalen" voor automatisch te doen
#kies juiste tekst uit mailberichtteksten.py
#worddocument omzetten naar html: gebruik bv. https://word2cleanhtml.com/
namenlijstcsv = "adressenlijst.csv"
htmlSjabloon = mailtemplate
#kies onderwerp en adres vanwaar mail moet worden verzonden (enkel nodig bij shared mailboxes)
mailSubject = ""
mailVerzendAdres = ""


df_rechten = pd.read_csv(namenlijstcsv)

for i in range(len(df_rechten)):
    #verzamel de benodigde variabelen uit de csv en datum + vul het template in met correcte adressering en achternaam
    firstname = df_rechten['firstname'].values[i]
    lastname = df_rechten['lastname'].values[i]
    title = df_rechten['Title'].values[i]
    padnaam = df_rechten['pathname'].values[i]
    objectenlijst = df_rechten["objectenlijst"].values[i]
    toestemming = df_rechten["toestemming"].values[i]
    mailadres = df_rechten['email'].values[i]
    vandaag = date.today()
    datum = vandaag.strftime("%Y%m%d")
    htmltext = Template(htmlSjabloon).safe_substitute(title=title, lastname=lastname, firstname=firstname)
    # open outlook, stel de mail op met onderwerp, bijlagen en correcte tekst en toon voor een final check
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = mailVerzendAdres
    mail.To = mailadres
    mail.Subject = mailSubject
    mail.GetInspector
    #voeg standaardmailhandtekening toe en het ingevulde sjabloon
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + htmltext + mail.HTMLbody[index + 1:]
    mail.Attachments.Add(objectenlijst)
    mail.Attachments.Add(toestemming)
    #toont de mail voor eventuele aanpassingen en controle, script loopt pas verder na verzenden/verwijderen mail
    mail.Display(True)
    #script pauzeert zodat de verzonden mail zeker in "verzonden berichten" verschijnt
    time.sleep(10)
    #zoekt de laatst verzonden mail en slaat die op in de padnaam met correcte naamgeving
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    verzondenItems = outlook.GetDefaultFolder(5)
    messages = verzondenItems.items
    message = messages.GetLast()
    message.SaveAs(os.path.join(padnaam, (datum + "_V1.msg")))
    print(lastname + " Done")
