# installer les librairies suivantes
## pandas
## xlrd

import pandas as pd

# chargement du fichier excel
df = pd.read_excel('Student_db.xlsx', sheetname='Feuil1')

# Extraction et création d'un fichier EXCEL contenant les liste des élèves ayant la moyenne
moyenne = df.loc[df['Moyenne'] >= 10]
moyenne.to_excel("moyenne.xlsx", sheet_name='Feuil1')

# Extraction et création d'un fichier EXCEL contenant les élèves ayant plus de 20 ans
plusde20ans = df.loc[df['Age'] > 20]
plusde20ans.to_excel("plusde20ans.xlsx", sheet_name='Feuil1')

# Extraction et création d'un fichier EXCEL contenant les statistiques globales de l’école

## calcul de la moyenne de l'école
moyenne_ecole = df.loc[:,"Moyenne"].mean()

## nombre de fille dans l'école
eleves = df['Sexe'].value_counts().to_dict()
frequence_boy = eleves['M']
frequence_girl = eleves['F']
## total des élèves
total = frequence_boy+frequence_girl

# pourcentage de garçons
pourcentage_boy = (frequence_boy*100)/total
# pourcentage de filles
pourcentage_girl = (frequence_girl*100)/total

## région qui a les meilleurs élèves
moyenne_max = df.loc[:,"Moyenne"].max()
entree_pour_la_moyenne_max = df.loc[df['Moyenne'] == moyenne_max]
region = df.at[13, 'Région']

# génération du fichier excel de statisque sur l"école
stats = pd.DataFrame([[moyenne_ecole, pourcentage_girl, pourcentage_boy, region]],
        index = None,
        columns = ['Moyenne ecole', 'Pourcentage de fille', 'Pourcentage de garçons', 'Meilleur région'])
stats.to_excel("stats.xlsx", index=False)

