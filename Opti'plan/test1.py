from dataclasses import dataclass
from typing import List, Dict, Optional

@dataclass
class Professor:
    id: int
    name: str
    rank: str  # 'Professeur', 'Docteur', 'Maître de Conférences'
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

professors = [
    Professor(1, "Prof. A", "Docteur", ["Génie logiciel", "Intelligence artificielle"], list(range(0, 40))),
    Professor(2, "Prof. B", "Docteur", ["Sécurité informatique"], list(range(0, 40))),
    Professor(3, "Prof. C", "Maître de Conférences", ["Internet et Multimédia"], list(range(0, 40))),
    Professor(4, "Prof. D", "Professeur", ["Systèmes embarqués et IoT"], list(range(0, 40))),
]

students = [
    Student(1, "Étudiant X", "Licence", "Génie logiciel", 1),
    Student(2, "Étudiant Y", "Master", "Sécurité Informatique", 2),
    Student(3, "Étudiant Z", "Licence", "Internet et Multimédia", 3),
]

rooms = [
    Room(1, "Salle 101"),
    Room(2, "Salle 102"),
]

# 8 créneaux par jour sur 5 jours
num_days = 5
slots_per_day = 8
total_slots = num_days * slots_per_day
time_slots = list(range(total_slots))

# Durée des soutenances : 1h pour la Licence, 1h30 pour le Master
level_durations = {'Licence': 1, 'Master': 1.5}

# Construction de l'emploi du temps
defenses: List[Defense] = []

# Étape 1 : Regroupement des créneaux en blocs de 4 heures
blocks = []
for day in range(num_days):
    for block_start in range(0, slots_per_day, 4):
        block_slots = [day * slots_per_day + block_start + i for i in range(4) if (block_start + i) < slots_per_day]
        blocks.append(block_slots)

# Fonction pour sélectionner un jury pour un bloc
def select_jury(professors: List[Professor], field: str) -> Optional[Dict[str, int]]:
    # Filtrer les professeurs éligibles pour être président
    president_candidates = [prof for prof in professors if prof.rank in ['Professeur', 'Docteur']]
    if not president_candidates:
        return None
    # Trier les présidents par rang (du plus élevé au plus bas)
    president_candidates.sort(key=lambda prof: {'Professeur': 3, 'Docteur': 2, 'Maître de Conférences': 1}[prof.rank], reverse=True)
    president = president_candidates[0]
    # Filtrer les examinateurs spécialistes du domaine
    examiner_candidates = [prof for prof in professors if field in prof.specialties and prof.id != president.id]
    if not examiner_candidates:
        return None
    examiner = examiner_candidates[0]
    return {'president_id': president.id, 'examiner_id': examiner.id}

# Initialisation des variables pour le suivi
scheduled_students = set()
professor_defense_count = {prof.id: 0 for prof in professors}
room_schedule = {room.id: [[] for _ in range(num_days)] for room in rooms}
professor_schedule = {prof.id: [] for prof in professors}

# Étape 2 : Planification des soutenances
for block_slots in blocks:
    day = block_slots[0] // slots_per_day
    available_students = [student for student in students if student.id not in scheduled_students]

    if not available_students:
        continue

    # Regrouper les étudiants par domaine
    field_students = {}
    for student in available_students:
        field_students.setdefault(student.field, []).append(student)

    # Pour chaque domaine, essayer d'affecter les étudiants
    for field, students_in_field in field_students.items():
        jury = select_jury(professors, field)
        if jury is None:
            continue

        president_id = jury['president_id']
        examiner_id = jury['examiner_id']

        # Vérifier que le président et l'examinateur n'ont pas dépassé 4 soutenances par jour
        if professor_defense_count[president_id] >= 4 or professor_defense_count[examiner_id] >= 4:
            continue

        # Planifier les étudiants dans le bloc
        for time_slot in block_slots:
            if not students_in_field:
                break

            student = students_in_field.pop(0)

            # Affecter une salle disponible
            room_assigned = None
            for room in rooms:
                room_day_schedule = room_schedule[room.id][day]
                if len(room_day_schedule) < 8 and time_slot not in room_day_schedule:
                    room_assigned = room
                    break
            if room_assigned is None:
                continue

            # Vérifier la disponibilité de l'encadreur
            supervisor = next((prof for prof in professors if prof.id == student.supervisor_id), None)
            if supervisor is None or time_slot not in supervisor.availability or time_slot in professor_schedule[supervisor.id]:
                continue

            # Vérifier que l'encadreur n'a pas dépassé 4 soutenances par jour
            if professor_defense_count[supervisor.id] >= 4:
                continue

            # Vérifier la disponibilité du président et de l'examinateur
            president = next((prof for prof in professors if prof.id == president_id), None)
            examiner = next((prof for prof in professors if prof.id == examiner_id), None)

            if time_slot not in president.availability or time_slot not in examiner.availability:
                continue

            if time_slot in professor_schedule[president.id] or time_slot in professor_schedule[examiner.id]:
                continue

            # Planifier la soutenance
            defense = Defense(
                student_id=student.id,
                time_slot=time_slot,
                room_id=room_assigned.id,
                president_id=president_id,
                examiner_id=examiner_id,
                supervisor_id=student.supervisor_id
            )
            defenses.append(defense)
            scheduled_students.add(student.id)
            professor_defense_count[president_id] += 1
            professor_defense_count[examiner_id] += 1
            professor_defense_count[supervisor.id] += 1

            # Mettre à jour les emplois du temps
            room_schedule[room_assigned.id][day].append(time_slot)
            professor_schedule[president.id].append(time_slot)
            professor_schedule[examiner.id].append(time_slot)
            professor_schedule[supervisor.id].append(time_slot)

# Affichage de l'emploi du temps
print("Emploi du temps des soutenances :")
for defense in defenses:
    student = next(s for s in students if s.id == defense.student_id)
    room = next(r for r in rooms if r.id == defense.room_id)
    president = next(p for p in professors if p.id == defense.president_id)
    examiner = next(p for p in professors if p.id == defense.examiner_id)
    supervisor = next(p for p in professors if p.id == defense.supervisor_id)
    time_slot = defense.time_slot
    day = time_slot // slots_per_day
    slot_in_day = time_slot % slots_per_day
    print(f"Jour {day+1}, Créneau {slot_in_day+1}: Étudiant {student.name} dans {room.name}")
    print(f"  Jury: Président {president.name}, Examinateur {examiner.name}, Encadreur {supervisor.name}\n")
