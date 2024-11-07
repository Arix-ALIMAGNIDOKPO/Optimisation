import pandas as pd

# Lire les données depuis le fichier Excel
df = pd.read_excel('Dispo_Enseignants (1).xlsx')

# Renommer les colonnes pour simplifier l'accès
df.rename(columns={
    'Nom et Prénom(s)': 'Enseignant',
    'Jour': 'Date',
    'Disponibilités (Heures)': 'Créneau horaire'
}, inplace=True)

# Convertir la colonne 'Date' en objet date
df['Date'] = pd.to_datetime(df['Date']).dt.date

# Liste des créneaux horaires à considérer (8 créneaux)
creneaux = [
    '08h00 - 09h00',
    '09h00 - 10H00',
    '10H00 - 11H00',
    '11H00 - 12H00',
    '13H00 - 14H00',
    '14H00 - 15H00',
    '15H00 - 16H00',
    '16H00 - 17H00'
]

# Création d'un mapping des créneaux horaires à des indices
creneau_indices = {creneau: idx for idx, creneau in enumerate(creneaux)}

# Extraire les dates uniques et les trier par ordre croissant
dates_uniques = sorted(df['Date'].unique())

# Obtenir la liste complète des enseignants (par exemple, à partir d'une liste fournie ou des données)
# Si vous avez une liste externe des enseignants, utilisez-la ici. Sinon, nous l'extrairons des données existantes.
enseignants_uniques = df['Enseignant'].unique()

# Initialiser un dictionnaire pour stocker les disponibilités
disponibilites_par_enseignant = {}

# Parcourir chaque enseignant
for enseignant in enseignants_uniques:
    # Initialiser la liste des disponibilités pour cet enseignant
    disponibilites_par_date = []
    # Parcourir chaque date triée
    for date in dates_uniques:
        # Filtrer les données pour cet enseignant et cette date
        df_enseignant_date = df[(df['Enseignant'] == enseignant) & (df['Date'] == date)]
        # Initialiser la liste des disponibilités pour les 8 créneaux à False
        dispo_jour = [False] * 8
        # Parcourir les créneaux horaires disponibles pour cet enseignant à cette date
        for creneau in df_enseignant_date['Créneau horaire']:
            # Vérifier si le créneau est dans notre liste de créneaux à considérer
            if creneau in creneau_indices:
                idx = creneau_indices[creneau]
                dispo_jour[idx] = True
        # Ajouter la liste des disponibilités pour ce jour
        disponibilites_par_date.append(dispo_jour)
    # Ajouter les disponibilités de l'enseignant dans le dictionnaire principal
    disponibilites_par_enseignant[enseignant] = disponibilites_par_date

# Maintenant, nous allons vérifier s'il y a des enseignants absents des données et les ajouter avec des disponibilités vides
# Si vous avez une liste complète des enseignants (par exemple, 'liste_complète_enseignants'), vous pouvez l'utiliser ici
# Par exemple :
# liste_complète_enseignants = ['Enseignant1', 'Enseignant2', 'Enseignant3', ..., 'EnseignantN']

# Pour cet exemple, supposons que tous les enseignants sont déjà dans 'enseignants_uniques'
# Si ce n'est pas le cas, vous pouvez décommenter et ajuster le code ci-dessous :

# for enseignant in liste_complète_enseignants:
#     if enseignant not in disponibilites_par_enseignant:
#         # Ajouter l'enseignant avec des disponibilités vides
#         disponibilites_par_enseignant[enseignant] = [[False]*8 for _ in dates_uniques]

# Préparer les données pour l'exportation
resultats = []

for enseignant in enseignants_uniques:
    disponibilites = disponibilites_par_enseignant.get(enseignant, [[False]*8 for _ in dates_uniques])
    for date_idx, dispo_jour in enumerate(disponibilites):
        resultats.append({
            'Enseignant': enseignant,
            'Date': dates_uniques[date_idx],
            'Disponibilités': dispo_jour
        })

# Convertir en DataFrame pour l'exportation
df_resultats = pd.DataFrame(resultats)

# Séparer les disponibilités en colonnes individuelles pour chaque créneau
for idx in range(8):
    df_resultats[f'Créneau {idx + 1}'] = df_resultats['Disponibilités'].apply(lambda x: x[idx])

# Supprimer la colonne 'Disponibilités' initiale
df_resultats.drop('Disponibilités', axis=1, inplace=True)

# Trier les données par enseignant et date
df_resultats.sort_values(by=['Enseignant', 'Date'], inplace=True)

# Exporter vers un nouveau fichier Excel
df_resultats.to_excel('disponibilites_par_enseignant_par_date.xlsx', index=False)
