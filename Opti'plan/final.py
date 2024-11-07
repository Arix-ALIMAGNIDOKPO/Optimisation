import pandas as pd
from openpyxl import Workbook

# Charger les données
file_path = 'disponibilites_par_enseignant_par_date.xlsx'
df = pd.read_excel(file_path)

# Spécifiez les dates dans l'ordre souhaité
specified_dates = [
    "2024-06-20", "2024-06-24", "2024-06-25", "2024-06-26", 
    "2024-06-27", "2024-06-28", "2024-07-01"
]
specified_dates_dt = pd.to_datetime(specified_dates)

# Filtrer les données pour les dates spécifiées
df['Date'] = pd.to_datetime(df['Date'])
df_filtered = df[df['Date'].isin(specified_dates_dt)]

# Créer le fichier Excel
wb = Workbook()
ws = wb.active
ws.title = "Disponibilités par Enseignant"
ws.append(["Enseignant", "Disponibilités"])

# Traiter chaque enseignant
for name, group in df_filtered.groupby('Enseignant'):
    daily_availabilities = []
    for date in specified_dates_dt:
        # Vérifier si l'enseignant est disponible à cette date
        row = group[group['Date'] == date]
        if not row.empty:
            # Obtenir la liste des créneaux pour cette date
            slots = row.iloc[0][['Créneau 1', 'Créneau 2', 'Créneau 3', 'Créneau 4', 
                                 'Créneau 5', 'Créneau 6', 'Créneau 7', 'Créneau 8']].tolist()
        else:
            # Si aucune donnée pour cette date, remplir avec False
            slots = [False] * 8
        daily_availabilities.append(slots)
    
    # Écrire dans la feuille Excel
    ws.append([name, str(daily_availabilities)])

# Sauvegarder le fichier
output_path = 'disponibilites_format_specifique.xlsx'
wb.save(output_path)
