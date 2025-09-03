import yaml
import os
from datetime import datetime

class Database:
    def __init__(self, filename='data/ecoles.yaml'):
        self.filename = filename
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        if not os.path.exists(filename):
            self.save_data({'ecoliers': [], 'eleves': [], 'notes': []})
    
    def load_data(self):
        try:
            with open(self.filename, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {'ecoliers': [], 'eleves': [], 'notes': []}
        except:
            return {'ecoliers': [], 'eleves': [], 'notes': []}
    
    def save_data(self, data):
        with open(self.filename, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True, default_flow_style=False)
    
    def add_ecolier(self, ecolier):
        data = self.load_data()
        ecolier['id'] = len(data['ecoliers']) + 1
        ecolier['date_inscription'] = datetime.now().strftime('%d/%m/%Y %H:%M')
        data['ecoliers'].append(ecolier)
        self.save_data(data)
        return ecolier['id']
    
    def add_eleve(self, eleve):
        data = self.load_data()
        eleve['id'] = len(data['eleves']) + 1
        eleve['date_inscription'] = datetime.now().strftime('%d/%m/%Y %H:%M')
        data['eleves'].append(eleve)
        self.save_data(data)
        return eleve['id']
    
    def get_ecoliers(self):
        return self.load_data()['ecoliers']
    
    def get_eleves(self):
        return self.load_data()['eleves']
    
    def get_all(self):
        data = self.load_data()
        all_students = []
        
        for ecolier in data['ecoliers']:
            ecolier['type'] = 'ecolier'
            all_students.append(ecolier)
        
        for eleve in data['eleves']:
            eleve['type'] = 'eleve'
            all_students.append(eleve)
        
        return all_students
    
    def update_ecolier(self, ecolier_id, updates):
        data = self.load_data()
        for ecolier in data['ecoliers']:
            if ecolier['id'] == ecolier_id:
                ecolier.update(updates)
                self.save_data(data)
                return True
        return False
    
    def update_eleve(self, eleve_id, updates):
        data = self.load_data()
        for eleve in data['eleves']:
            if eleve['id'] == eleve_id:
                eleve.update(updates)
                self.save_data(data)
                return True
        return False
    
    def add_payment(self, student_id, student_type, amount):
        data = self.load_data()
        
        if student_type == 'ecolier':
            for student in data['ecoliers']:
                if student['id'] == student_id:
                    if 'payments' not in student:
                        student['payments'] = []
                    student['payments'].append({
                        'date': datetime.now().strftime('%d/%m/%Y %H:%M'),
                        'amount': amount
                    })
                    self.save_data(data)
                    return True
        
        elif student_type == 'eleve':
            for student in data['eleves']:
                if student['id'] == student_id:
                    if 'payments' not in student:
                        student['payments'] = []
                    student['payments'].append({
                        'date': datetime.now().strftime('%d/%m/%Y %H:%M'),
                        'amount': amount
                    })
                    self.save_data(data)
                    return True
        
        return False
    
    def get_total_paid(self, student):
        total = 0
        if 'payments' in student:
            for payment in student['payments']:
                total += int(payment['amount'])
        return total
    
    def add_note(self, student_id, student_type, classe, matiere, note):
        data = self.load_data()
        
        note_entry = {
            'student_id': student_id,
            'student_type': student_type,
            'classe': classe,
            'matiere': matiere,
            'note': note,
            'date': datetime.now().strftime('%d/%m/%Y %H:%M')
        }
        
        data['notes'].append(note_entry)
        self.save_data(data)
        return True
    
    def get_notes(self, classe=None, matiere=None):
        data = self.load_data()
        notes = data['notes']
        
        if classe:
            notes = [n for n in notes if n['classe'] == classe]
        if matiere:
            notes = [n for n in notes if n['matiere'] == matiere]
        
        return notes
    
    def get_student_notes(self, student_id, student_type):
        notes = self.get_notes()
        return [n for n in notes if n['student_id'] == student_id and n['student_type'] == student_type]
