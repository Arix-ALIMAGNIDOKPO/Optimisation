from dataclasses import dataclass
from typing import List, Dict, Optional
import random
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd

@dataclass
class Professor:
    id: int
    name: str
    rank: str  # 'MC', 'Docteur', 'Professeur'
    specialties: List[str]
    availability: List[int]  # Liste des créneaux disponibles

@dataclass
class Student:
    id: int
    name: str
    level: str  # 'Licence' ou 'Master'
    field: str
    supervisor_id: int

@dataclass
class Room:
    id: int
    name: str

@dataclass
class Defense:
    student_id: int
    time_slot: int
    room_id: int
    president_id: int
    examiner_id: int
    supervisor_id: int

"""
# Génération des données pour les professeurs
random.seed(42)  # Pour des résultats reproductibles
professor_names = [
    "Enseignant " + chr(65 + i) for i in range(10)
]
ranks = ['MC', 'Docteur', 'Professeur']
rank_values = {'MC': 3, 'Docteur': 2, 'Professeur': 1}
fields = ['Génie logiciel', 'Intelligence artificielle', 'Sécurité informatique',
          'Internet et Multimédia', 'Systèmes embarqués et IoT']

professors = []
for i in range(10):
    prof_id = i + 1
    name = professor_names[i]
    rank = random.choice(ranks)
    specialties = random.sample(fields, k=random.randint(1, 2))
    availability = list(range(0, 40))  # Supposons qu'ils sont disponibles sur tous les créneaux
    professors.append(Professor(prof_id, name, rank, specialties, availability))
"""
# Lecture du fichier Excel
df_professors = pd.read_excel('professors_data.xlsx')

# Définition des grades acceptés dans votre code
grade_mapping = {
    'Docteur': 'Docteur',
    'Ingénieur': 'Professeur',
    'Professeur': 'Professeur',
    'MA': 'MC',
    'MC': 'MC'
}

# Liste pour stocker les objets Professor
professors = []
for index, row in df_professors.iterrows():
    prof_id = row['Numéro']
    name = f"{row['Nom']} {row['Prénoms']}"
    grade = row['Grade']

    # Conversion du grade en utilisant le mapping
    rank = grade_mapping.get(grade, 'Professeur')  # Par défaut, considérer comme 'Professeur' si non trouvé

    # Traitement des disponibilités
    availability = []
    raw_availability = eval(row['Disponibilité'])  # Utiliser `eval` pour convertir la chaîne en liste de listes
    for day_index, day_slots in enumerate(raw_availability):
        for slot_index, is_available in enumerate(day_slots):
            if is_available:  # Si la valeur est `True`, ajouter l'indice du créneau
                availability.append(day_index * 8 + slot_index)

    # Traitement des spécialités
    specialties = [s.strip() for s in row['Speciality'].split(',')]

    # Création de l'objet Professor
    professors.append(Professor(prof_id, name, rank, specialties, availability))

# Affichage des données des professeurs pour vérification
for prof in professors:
    print(prof)

"""
# Génération des données pour les étudiants
student_names = [
    "Étudiant " + str(i + 1) for i in range(100)
]
levels = ['Licence', 'Master']

students = []
for i in range(100):
    student_id = i + 1
    name = student_names[i]
    level = random.choice(levels)
    field = random.choice(fields)
    supervisor = random.choice(professors)
    students.append(Student(student_id, name, level, field, supervisor.id))
"""

# Nouveau code pour lire les étudiants à partir du fichier Excel
df_students = pd.read_excel('students_data.xlsx')

students = []
for index, row in df_students.iterrows():
    student_id = row['Numéro']
    name = row['Nom']
    level = row['Cycle']
    field = row['Filière']
    supervisor_id = row['MM']  # Assurez-vous que cette colonne contient l'ID de l'encadreur

    students.append(Student(student_id, name, level, field, supervisor_id))

# Génération des salles
rooms = []
for i in range(5):
    room_id = i + 1
    name = f"Salle {100 + i}"
    rooms.append(Room(room_id, name))

# Définition des créneaux (8 créneaux par jour sur 5 jours)
num_days = 5
slots_per_day = 8
total_slots = num_days * slots_per_day
time_slots = list(range(total_slots))

# Construction de l'emploi du temps
defenses: List[Defense] = []

# Initialisation des variables pour le suivi
scheduled_students = set()
professor_defense_count = {prof.id: {day: 0 for day in range(num_days)} for prof in professors}
professor_total_defenses = {prof.id: 0 for prof in professors}
room_schedule = {room.id: {slot: None for slot in time_slots} for room in rooms}
professor_schedule = {prof.id: set() for prof in professors}

# Trier les professeurs par grade décroissant
professors_sorted = sorted(professors, key=lambda x: rank_values[x.rank], reverse=True)

# Pré-calculer les examinateurs possibles pour chaque étudiant
student_examiner_options = {}
for student in students:
    examiners = [prof for prof in professors if student.field in prof.specialties and prof.id != student.supervisor_id]
    student_examiner_options[student.id] = examiners

# Trier les étudiants par nombre croissant d'examinateurs possibles (priorité aux étudiants avec moins d'options)
students_sorted = sorted(students, key=lambda x: len(student_examiner_options[x.id]))

# Planification des soutenances
student_index = 0
for slot in time_slots:
    # Récupérer le jour actuel
    day = slot // slots_per_day

    for room in rooms:
        if room_schedule[room.id][slot] is not None:
            continue  # Créneau déjà utilisé

        # Essayer de planifier un étudiant dans ce créneau
        for i in range(len(students_sorted)):
            student = students_sorted[(student_index + i) % len(students_sorted)]
            if student.id in scheduled_students:
                continue  # Étudiant déjà programmé

            supervisor = next((prof for prof in professors if prof.id == student.supervisor_id), None)
            if supervisor is None:
                continue

            if slot not in supervisor.availability or slot in professor_schedule[supervisor.id]:
                continue

            if professor_defense_count[supervisor.id][day] >= 4:
                continue

            # Déterminer le grade minimal requis pour le président en fonction du niveau de l'étudiant
            if student.level == 'Licence':
                min_president_rank_value = rank_values['Docteur']
            else:  # Master
                min_president_rank_value = rank_values['MC']

            # Le président doit avoir un grade supérieur ou égal à celui de l'encadreur
            supervisor_rank_value = rank_values[supervisor.rank]
            min_president_rank_value = max(min_president_rank_value, supervisor_rank_value)

            # Le président ne peut pas être de grade 'Professeur'
            eligible_presidents = [prof for prof in professors_sorted
                                   if prof.rank != 'Professeur' and
                                   rank_values[prof.rank] >= min_president_rank_value and
                                   prof.id not in [supervisor.id] and
                                   slot in prof.availability and
                                   slot not in professor_schedule[prof.id] and
                                   professor_defense_count[prof.id][day] < 4]

            if not eligible_presidents:
                continue
            president = eligible_presidents[0]  # Le plus gradé disponible

            # Sélectionner un examinateur spécialiste du domaine, différent du président et de l'encadreur
            examiners = [prof for prof in student_examiner_options[student.id]
                         if prof.id not in [supervisor.id, president.id] and
                         slot in prof.availability and
                         slot not in professor_schedule[prof.id] and
                         professor_defense_count[prof.id][day] < 4]
            if not examiners:
                continue
            examiner = examiners[0]  # Choisir le premier disponible

            # Planifier la soutenance
            defense = Defense(
                student_id=student.id,
                time_slot=slot,
                room_id=room.id,
                president_id=president.id,
                examiner_id=examiner.id,
                supervisor_id=supervisor.id
            )
            defenses.append(defense)
            scheduled_students.add(student.id)

            # Mettre à jour les emplois du temps
            room_schedule[room.id][slot] = defense
            professor_schedule[supervisor.id].add(slot)
            professor_schedule[president.id].add(slot)
            professor_schedule[examiner.id].add(slot)

            professor_defense_count[supervisor.id][day] += 1
            professor_defense_count[president.id][day] += 1
            professor_defense_count[examiner.id][day] += 1

            professor_total_defenses[supervisor.id] += 1
            professor_total_defenses[president.id] += 1
            professor_total_defenses[examiner.id] += 1

            student_index = (student_index + 1) % len(students_sorted)
            break  # Passer au créneau suivant après avoir planifié une soutenance
        else:
            # Aucun étudiant n'a pu être programmé dans ce créneau
            continue

# Vérifier si tous les créneaux ont été utilisés
unused_slots = [slot for slot in time_slots for room in rooms if room_schedule[room.id][slot] is None]
if unused_slots:
    # Si des créneaux sont encore disponibles et qu'il reste des étudiants non programmés, essayer de les affecter
    for slot in unused_slots:
        day = slot // slots_per_day
        for room in rooms:
            if room_schedule[room.id][slot] is not None:
                continue  # Créneau déjà utilisé

            # Essayer de planifier un étudiant non programmé
            for student in students_sorted:
                if student.id in scheduled_students:
                    continue  # Étudiant déjà programmé

                supervisor = next((prof for prof in professors if prof.id == student.supervisor_id), None)
                if supervisor is None:
                    continue

                if slot not in supervisor.availability or slot in professor_schedule[supervisor.id]:
                    continue

                if professor_defense_count[supervisor.id][day] >= 4:
                    continue

                # Déterminer le grade minimal requis pour le président en fonction du niveau de l'étudiant
                if student.level == 'Licence':
                    min_president_rank_value = rank_values['Docteur']
                else:  # Master
                    min_president_rank_value = rank_values['MC']

                # Le président doit avoir un grade supérieur ou égal à celui de l'encadreur
                supervisor_rank_value = rank_values[supervisor.rank]
                min_president_rank_value = max(min_president_rank_value, supervisor_rank_value)

                # Le président ne peut pas être de grade 'Professeur'
                eligible_presidents = [prof for prof in professors_sorted
                                       if prof.rank != 'Professeur' and
                                       rank_values[prof.rank] >= min_president_rank_value and
                                       prof.id not in [supervisor.id] and
                                       slot in prof.availability and
                                       slot not in professor_schedule[prof.id] and
                                       professor_defense_count[prof.id][day] < 4]

                if not eligible_presidents:
                    continue
                president = eligible_presidents[0]  # Le plus gradé disponible

                # Sélectionner un examinateur, même s'il n'est pas spécialiste du domaine
                examiners = [prof for prof in professors
                             if prof.id not in [supervisor.id, president.id] and
                             slot in prof.availability and
                             slot not in professor_schedule[prof.id] and
                             professor_defense_count[prof.id][day] < 4]
                if not examiners:
                    continue
                examiner = examiners[0]

                # Planifier la soutenance
                defense = Defense(
                    student_id=student.id,
                    time_slot=slot,
                    room_id=room.id,
                    president_id=president.id,
                    examiner_id=examiner.id,
                    supervisor_id=supervisor.id
                )
                defenses.append(defense)
                scheduled_students.add(student.id)

                # Mettre à jour les emplois du temps
                room_schedule[room.id][slot] = defense
                professor_schedule[supervisor.id].add(slot)
                professor_schedule[president.id].add(slot)
                professor_schedule[examiner.id].add(slot)

                professor_defense_count[supervisor.id][day] += 1
                professor_defense_count[president.id][day] += 1
                professor_defense_count[examiner.id][day] += 1

                professor_total_defenses[supervisor.id] += 1
                professor_total_defenses[president.id] += 1
                professor_total_defenses[examiner.id] += 1

                break  # Passer au créneau suivant après avoir planifié une soutenance

# Collecte des étudiants non programmés
unscheduled_students = [student for student in students if student.id not in scheduled_students]

# Collecte des créneaux disponibles dans les salles
available_room_slots = {}
for room in rooms:
    available_slots = [slot for slot in time_slots if room_schedule[room.id][slot] is None]
    if available_slots:
        available_room_slots[room.name] = {}
        for day in range(num_days):
            day_slots = [slot for slot in available_slots if slot // slots_per_day == day]
            if day_slots:
                slots_in_day = [slot % slots_per_day + 1 for slot in day_slots]
                available_room_slots[room.name][f"Jour {day+1}"] = sorted(slots_in_day)

# Calcul des statistiques
total_students = len(students)
scheduled_students_count = len(scheduled_students)
unscheduled_students_count = total_students - scheduled_students_count
room_capacity = total_slots * len(rooms)
used_slots_count = len([slot for room in rooms for slot in time_slots if room_schedule[room.id][slot] is not None])
available_slots_count = room_capacity - used_slots_count
room_utilization = (used_slots_count / room_capacity) * 100

# Préparation du PDF
pdf_file = "planning_soutenances.pdf"
pdf = SimpleDocTemplate(pdf_file, pagesize=A4)
elements = []
styles = getSampleStyleSheet()
title_style = styles['Heading1']
subtitle_style = styles['Heading2']
normal_style = styles['Normal']

# Section 1: Planning des soutenances
elements.append(Paragraph("Planning des soutenances", title_style))
elements.append(Spacer(1, 12))

# Création du tableau pour le planning
data = [['Jour', 'Créneau', 'Étudiant', 'Niveau', 'Domaine', 'Salle', 'Président', 'Examinateur', 'Encadreur']]

for defense in defenses:
    student = next(s for s in students if s.id == defense.student_id)
    room = next(r for r in rooms if r.id == defense.room_id)
    president = next(p for p in professors if p.id == defense.president_id)
    examiner = next(p for p in professors if p.id == defense.examiner_id)
    supervisor = next(p for p in professors if p.id == defense.supervisor_id)
    time_slot = defense.time_slot
    day = time_slot // slots_per_day + 1  # Les jours commencent à 1
    slot_in_day = time_slot % slots_per_day + 1  # Les créneaux commencent à 1

    # Inclure le grade de chaque enseignant
    president_info = f"{president.name} ({president.rank})"
    examiner_info = f"{examiner.name} ({examiner.rank})"
    supervisor_info = f"{supervisor.name} ({supervisor.rank})"

    data.append([
        f"Jour {day}",
        f"Créneau {slot_in_day}",
        student.name,
        student.level,
        student.field,
        room.name,
        president_info,
        examiner_info,
        supervisor_info
    ])

table = Table(data, repeatRows=1)

# Style du tableau
style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # En-tête en gris
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Texte de l'en-tête en blanc
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Police en gras pour l'en-tête
    ('FONTSIZE', (0, 0), (-1, -1), 8),  # Taille de police réduite pour tout le tableau
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  # Grille noire
])

table.setStyle(style)

# Ajustement de la taille des colonnes si nécessaire
col_widths = [50, 50, 70, 50, 80, 50, 100, 100, 100]
table._argW = col_widths

# Alternance des couleurs pour les lignes
for i in range(1, len(data)):
    if i % 2 == 0:
        bg_color = colors.lightgrey
    else:
        bg_color = colors.white
    table.setStyle(TableStyle([('BACKGROUND', (0, i), (-1, i), bg_color)]))

elements.append(table)
elements.append(Spacer(1, 24))

# Section 2: Étudiants non programmés
elements.append(Paragraph("Étudiants non programmés", subtitle_style))
elements.append(Spacer(1, 12))

if unscheduled_students:
    data_unscheduled = [['ID Étudiant', 'Nom', 'Niveau', 'Domaine', 'Encadreur']]
    for student in unscheduled_students:
        supervisor = next(p for p in professors if p.id == student.supervisor_id)
        supervisor_info = f"{supervisor.name} ({supervisor.rank})"
        data_unscheduled.append([
            f"{student.id}",
            student.name,
            student.level,
            student.field,
            supervisor_info
        ])
    table_unscheduled = Table(data_unscheduled, repeatRows=1)

    # Style du tableau
    style_unscheduled = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])

    table_unscheduled.setStyle(style_unscheduled)
    elements.append(table_unscheduled)
else:
    elements.append(Paragraph("Tous les étudiants ont été programmés.", normal_style))

elements.append(Spacer(1, 24))

# Section 3: Créneaux disponibles dans les salles
elements.append(Paragraph("Créneaux disponibles dans les salles", subtitle_style))
elements.append(Spacer(1, 12))

data_rooms = [['Salle', 'Jour', 'Créneaux disponibles']]
for room_name, days in available_room_slots.items():
    for day, slots in days.items():
        slots_str = ', '.join([str(slot) for slot in slots])
        data_rooms.append([room_name, day, slots_str])

if len(data_rooms) > 1:
    table_rooms = Table(data_rooms, repeatRows=1)

    # Style du tableau
    style_rooms = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])

    table_rooms.setStyle(style_rooms)
    elements.append(table_rooms)
else:
    elements.append(Paragraph("Aucun créneau disponible dans les salles.", normal_style))

elements.append(Spacer(1, 24))

# Section 4: Statistiques sur l'efficacité de l'algorithme
elements.append(Paragraph("Statistiques sur l'efficacité de l'algorithme", subtitle_style))
elements.append(Spacer(1, 12))

stats_paragraph = Paragraph(f"""
Nombre total d'étudiants : {total_students}<br/>
Nombre d'étudiants programmés : {scheduled_students_count}<br/>
Nombre d'étudiants non programmés : {unscheduled_students_count}<br/>
Pourcentage d'étudiants programmés : {scheduled_students_count / total_students * 100:.2f}%<br/>
Nombre total de créneaux disponibles : {room_capacity}<br/>
Nombre de créneaux utilisés : {used_slots_count}<br/>
Nombre de créneaux restants : {available_slots_count}<br/>
Pourcentage d'utilisation des salles : {room_utilization:.2f}%<br/>
""", normal_style)

elements.append(stats_paragraph)

# Génération du PDF
pdf.build(elements)

print(f"Le planning des soutenances a été généré dans le fichier '{pdf_file}'.")
