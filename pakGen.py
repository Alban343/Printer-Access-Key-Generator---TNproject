import pandas as pd
import numpy as np
# Les colonnes dans l'ordre correspondant au format du csv importable dans l'utilitaire Ricoh
columns = ['<type>', '<name>', '<displayName>', '<phoneticName>', '<common>',
       '<tagSet1>', '<tagSet2>', '<tagSet3>', '<address>', '<isSender>',
       '<protect>', '<password>', '<userCode>', '<group>', '<faxNumber>',
       '<lineType>', '<isAbroad>', '<ttiNo>', '<label1>', '<label2String>',
       '<messageNo>', '<protectFolder>', '<passwordEncoding>',
       '<folderProtocol>', '<ftpPort>', '<folderServer>', '<folderPath>',
       '<folderUser>', '<ftpCharCoding>', '<entryACL>', '<documentACL>',
       '<IPfaxProtocol>', '<IPfaxAddress>', '<authPassword>',
       '<passwordEncoding2>', '<SMTPAuth>', '<SMTPUser>', '<SMTPPassword>',
       '<passwordEncoding3>', '<folderAuth>', '<folderPassword>',
       '<passwordEncoding4>', '<LDAPAuth>', '<LDAPUser>', '<LDAPPassword>',
       '<passwordEncoding5>', '<DirectSMTP>']

dictionnary = dict.fromkeys(columns)
dictionnary['<type>'] = []
classeur = pd.DataFrame.from_dict(dictionnary)

# Importation et mise au format de l'annuaire depuis un fichier csv
annuaire_df = pd.read_csv('Annuaire.csv', encoding='utf-8')
annuaire_df['<address>'] = annuaire_df['mail']
annuaire_df['<name>'] = annuaire_df['prenom'] + ' ' + annuaire_df['nom'].apply(lambda x: str.upper(x))
annuaire_df['<displayName>'] = annuaire_df['prenom'] + ' ' + annuaire_df['nom']

add_temp = annuaire_df[['<name>', '<displayName>', '<address>']]

classeur_export = pd.merge(add_temp, classeur, on=['<name>', '<displayName>', '<address>'], how='outer')

# Formattage des options récurrentes
classeur_export['<type>'] = '[A]'
classeur_export[['<passwordEncoding3>', '<passwordEncoding4>', '<passwordEncoding5>']] = '[omitted]'
classeur_export[['<password>', '<group>', '<passwordEncoding>', '<ftpPort>', '<folderServer>', '<folderUser>', '<ftpCharCoding>', '<SMTPUser>', '<SMTPPassword>', '<folderPassword>', '<LDAPUser>', '<LDAPPassword>']] = '[]'
classeur_export[['<tagSet1>', '<tagSet2>', '<tagSet3>', '<protect>', '<protectFolder>', '<folderProtocol>', '<SMTPAuth>', '<folderAuth>', '<LDAPAuth>']] = '[0]'
classeur_export[['<common>', '<isSender>']] = '[1]'
classeur_export['<folderPath>'] = classeur_export['<name>'].apply(lambda x: f'[\\\\RV347727-DATA\\Scanner$\\{x}]')

# Création de codes d'accès personnalisés pour chaque utilisateur
classeur_export['<userCode>'] = np.random.randint(1000, 10000, size=len(classeur_export))

# Fonction de formatage pour compatibilité avec l'utilistaire Ricoh
def crochets(x):
    if isinstance(x, str):
       return f'[{x}]'
    else:
        return f'[{x}]'

cols_crochets = ['<name>', '<displayName>', '<address>', '<userCode>']
for col in cols_crochets:
    classeur_export[col] = classeur_export[col].apply(crochets)

# Index puis finalisation du formatage
classeur_export['<index>'] = [f'[{i:05d}]' for i in range(1, len(classeur_export) + 1)]
classeur_export = classeur_export.set_index('<index>')
classeur_export = classeur_export[columns]

# Export
classeur_export.to_csv("Printer_Logs.csv", encoding='latin1')