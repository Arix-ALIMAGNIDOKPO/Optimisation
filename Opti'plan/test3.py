import pandas as pd
import random
from dataclasses import dataclass
from typing import List, Dict, Optional
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

@dataclass
class Professor:
    id: int
    name: str
    rank: str  # 'MC', 'Docteur', 'Professeur'
    specialties: List[str]
    availability: List[int]

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

# Définition des grades et mapping
grade_mapping = {
    'Docteur': 'Docteur',
    'Ingénieur': 'Professeur',
    'Professeur': 'Professeur',
    'MA': 'MC',
    'MC': 'MC'
}

# Lecture des données des enseignants depuis l'Excel
df_professors = pd.read_excel('fichiers/enseignants.xlsx')
professors = []
for index, row in df_professors.iterrows():
    prof_id = row['Numéro']
    name = f"{row['Nom']} {row['Prénoms']}"
    grade = row['Grade']
    rank = grade_mapping.get(grade, 'Professeur')

    # Traitement des disponibilités
    availability = []
    raw_availability = eval(row['Disponibilité'])
    for day_index, day_slots in enumerate(raw_availability):
        for slot_index, is_available in enumerate(day_slots):
            if is_available:
                availability.append(day_index * 8 + slot_index)

    # Traitement des spécialités
    specialties = [s.strip() for s in row['Speciality'].split(',')]

    professors.append(Professor(prof_id, name, rank, specialties, availability))

# Lecture des données des étudiants depuis l'Excel
df_students = pd.read_excel('fichiers/students_data.xlsx')
students = []
for index, row in df_students.iterrows():
    student_id = row['Numéro']
    name = row['Nom']
    level = row['Cycle']
    field = row['Filière']
    supervisor_id = row['MM']
    students.append(Student(student_id, name, level, field, supervisor_id))

# Génération des salles
# Liste des salles
room_names = [
    'Zone Master A2-1', 
    'Zone Master A2-2',
    'Batiment ISA/FSA', 
    'Batiment RESBIO/FSA', 
    'Batiment SOKPON 1'
]

# Création des objets Room
rooms = [Room(id=i + 1, name=room_names[i]) for i in range(len(room_names))]

# Définition des créneaux (8 créneaux par jour sur 5 jours)
num_days = 5
slots_per_day = 8
total_slots = num_days * slots_per_day
time_slots = list(range(total_slots))

# Construction de l'emploi du temps
defenses: List[Defense] = []
scheduled_students = set()
rank_values = {'MC': 3, 'Docteur': 2, 'Professeur': 1}

# Préparation du tri des professeurs
professors_sorted = sorted(professors, key=lambda x: rank_values[x.rank], reverse=True)

# Pré-calcul des examinateurs possibles pour chaque étudiant
student_examiner_options = {}
for student in students:
    examiners = [prof for prof in professors if student.field in prof.specialties and prof.id != student.supervisor_id]
    student_examiner_options[student.id] = examiners

# Planification des soutenances
students_sorted = sorted(students, key=lambda x: len(student_examiner_options[x.id]))
professor_schedule = {prof.id: set() for prof in professors}
room_schedule = {room.id: {slot: None for slot in time_slots} for room in rooms}

for student in students_sorted:
    supervisor = next((prof for prof in professors if prof.id == student.supervisor_id), None)
    if supervisor is None:
        continue

    for slot in time_slots:
        if slot not in supervisor.availability or slot in professor_schedule[supervisor.id]:
            continue

        eligible_presidents = [prof for prof in professors_sorted if
                               prof.rank != 'Professeur' and
                               prof.id != supervisor.id and
                               slot in prof.availability and
                               slot not in professor_schedule[prof.id]]
        if not eligible_presidents:
            continue
        president = eligible_presidents[0]

        examiners = [prof for prof in student_examiner_options[student.id] if
                     prof.id not in [supervisor.id, president.id] and
                     slot in prof.availability and slot not in professor_schedule[prof.id]]
        if not examiners:
            continue
        examiner = examiners[0]

        room = next((r for r in rooms if room_schedule[r.id][slot] is None), None)
        if room is None:
            continue

        defense = Defense(student.id, slot, room.id, president.id, examiner.id, supervisor.id)
        defenses.append(defense)
        scheduled_students.add(student.id)
        room_schedule[room.id][slot] = defense
        professor_schedule[supervisor.id].add(slot)
        professor_schedule[president.id].add(slot)
        professor_schedule[examiner.id].add(slot)
        break

# Génération du PDF
pdf_file = "Soutenance.pdf"
pdf = SimpleDocTemplate(pdf_file, pagesize=A4)
elements = []
styles = getSampleStyleSheet()

elements.append(Paragraph("Planning des soutenances", styles['Heading1']))
data = [['Jour', 'Créneau', 'Étudiant', 'Niveau', 'Domaine', 'Salle', 'Président', 'Examinateur', 'Encadreur']]

for defense in defenses:
    student = next(s for s in students if s.id == defense.student_id)
    room = next(r for r in rooms if r.id == defense.room_id)
    president = next(p for p in professors if p.id == defense.president_id)
    examiner = next(p for p in professors if p.id == defense.examiner_id)
    supervisor = next(p for p in professors if p.id == defense.supervisor_id)
    day = defense.time_slot // slots_per_day + 1
    slot_in_day = defense.time_slot % slots_per_day + 1
    data.append([
        f"Jour {day}",
        f"Créneau {slot_in_day}",
        student.name,
        student.level,
        student.field,
        room.name,
        f"{president.name} ({president.rank})",
        f"{examiner.name} ({examiner.rank})",
        f"{supervisor.name} ({supervisor.rank})"
    ])

table = Table(data)
table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
]))

elements.append(table)
pdf.build(elements)

print("Le fichier PDF 'Soutenance.pdf' a été généré avec succès.")
