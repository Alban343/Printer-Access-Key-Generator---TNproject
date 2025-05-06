import pandas as pd
import smtplib
from email.message import EmailMessage

# Importation du csv et sélection des lignes pertinentes pour le mail
classeur_full = pd.read_csv("Printer_Logs.csv", encoding='latin1')
classeur = classeur_full[['<name>', '<address>', '<userCode>']]

# Définition d'alias afin de fluidifier les messages
alias = {"DGS" : "M. Carpentier",
         "Maire" : "M. Vasseur",
         "Police Resp" : "Silvère",
         "Gestion Personnel" : "Ysaline",
         "Police" : "Thibert",
         "CCAS" : "Azalaïs",
         "Sports-Vacances Act." : "Angilberte",
         "Espace Jeunes" : "Didier",
         "Cuisine centrale" : "Camillien",
         "Multi-Accueil" : "Sophie",
         "RIPAM" : "Lambertine",
         "Médiathèque" : "Aurèle",
         "CTM BAT" : "Gontran",
         "Vie Associative" : "Alcidie"}

# ----- Fonctions appelées par la fonction d'envoi de mail -----
# Définition du destinataire
def destinataire(browse):
    name = ''
    mail = f'{classeur.iloc[browse, 1]}'
    if classeur.iloc[browse, 0] in alias.keys():
        name = alias[f'{classeur.iloc[browse, 0]}']
    else:
        name = classeur.iloc[browse, 0]
    
    return([name, mail])
# Le code correspondant à transmettre
def printer_log(browse):
    log = f'{classeur.iloc[browse, 2]}'
    return f'{log}'
# Le message à éditer
def message(browse):
    subject = f'Nouveaux codes d\'imprimantes - {destinataire(browse)[0]}'
    message = (f'Bonjour {destinataire(browse)[0]}!\n'
               f'Voici le nouvel ID utilisateur pour les imprimantes : {printer_log(browse)}\n'
               f'Ceci est un code personnel.\n'
               f'Bonne journée !\n'
               f'Le service informatique')
    return([subject, message])
# Fonction d'envoi du message
def send_message(browse):

    smtp_server = "smtp.office365.com"
    smtp_port = 587
    email_address = "infra@ailleurs-mairie.fr"
    email_password = "gfdR$gdsnl*dne))Lmsnds^fbe"

    msg = EmailMessage()
    msg["Subject"] = message(browse)[0]
    msg["From"] = email_address
    msg["To"] = destinataire(browse)[1]
    msg.set_content(message(browse)[1])

    file_path = "Tutoriel ID utilisateur (imprimantes).pdf"
    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = f.name.split("/")[-1]

        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.login(email_address, email_password)
            server.send_message(msg)
            print("Email envoyé avec succès !")
    except Exception as e:
        print(f"Erreur lors de l'envoi : {e}")
# Envoi de message automatisé depuis le fichier exportable directement depuis les imprimantes Ricoh
def auto_message():
    for i in classeur.index:
        send_message(i)

#auto_message()