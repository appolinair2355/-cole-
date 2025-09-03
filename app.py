from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file
from models import Database
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io

app = Flask(__name__)
db = Database()

@app.route('/')
def accueil():
    return render_template('accueil.html')

@app.route('/inscription')
def inscription():
    return render_template('inscription.html')

@app.route('/inscrire_ecolier', methods=['POST'])
def inscrire_ecolier():
    data = request.json
    ecolier = {
        'nom': data['nom'],
        'prenoms': data['prenoms'],
        'sexe': data['sexe'],
        'date_naissance': data['date_naissance'],
        'classe': data['classe'],
        'numero_parents': data['numero_parents'],
        'montant_scolarite': data['montant_scolarite'],
        'nom_enregistreur': data['nom_enregistreur']
    }
    
    ecolier_id = db.add_ecolier(ecolier)
    return jsonify({'success': True, 'id': ecolier_id})

@app.route('/inscrire_eleve', methods=['POST'])
def inscrire_eleve():
    data = request.json
    eleve = {
        'nom': data['nom'],
        'prenoms': data['prenoms'],
        'sexe': data['sexe'],
        'date_naissance': data['date_naissance'],
        'classe': data['classe'],
        'numero_parents': data['numero_parents'],
        'montant_scolarite': data['montant_scolarite'],
        'nom_enregistreur': data['nom_enregistreur']
    }
    
    eleve_id = db.add_eleve(eleve)
    return jsonify({'success': True, 'id': eleve_id})

@app.route('/liste_eleves')
def liste_eleves():
    eleves = db.get_eleves()
    return render_template('liste_eleves.html', eleves=eleves)

@app.route('/liste_ecoliers')
def liste_ecoliers():
    ecoliers = db.get_ecoliers()
    return render_template('liste_ecoliers.html', ecoliers=ecoliers)

@app.route('/scolarite')
def scolarite():
    students = db.get_all()
    for student in students:
        student['total_paid'] = db.get_total_paid(student)
        student['reste'] = int(student['montant_scolarite']) - student['total_paid']
    return render_template('scolarite.html', students=students)

@app.route('/paiement', methods=['POST'])
def paiement():
    data = request.json
    success = db.add_payment(
        data['student_id'],
        data['student_type'],
        data['amount']
    )
    return jsonify({'success': success})

@app.route('/notes')
def notes():
    ecoliers = db.get_ecoliers()
    eleves = db.get_eleves()
    
    classes_ecoliers = list(set([e['classe'] for e in ecoliers]))
    classes_eleves = list(set([e['classe'] for e in eleves]))
    
    matieres = {
        'maternelle': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        'CI': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        'CP': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        'CE1': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        'CE2': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        'CM1': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        'CM2': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        '6ième': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        '5ième': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        '4ième': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol'],
        '3ième': ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol']
    }
    
    return render_template('notes.html', 
                         classes_ecoliers=classes_ecoliers,
                         classes_eleves=classes_eleves,
                         matieres=matieres)

@app.route('/get_students_by_class', methods=['POST'])
def get_students_by_class():
    data = request.json
    classe = data['classe']
    is_ecolier = data['is_ecolier']
    
    if is_ecolier:
        students = [s for s in db.get_ecoliers() if s['classe'] == classe]
    else:
        students = [s for s in db.get_eleves() if s['classe'] == classe]
    
    return jsonify({'students': students})

@app.route('/save_notes', methods=['POST'])
def save_notes():
    data = request.json
    notes = data['notes']
    
    for note_data in notes:
        db.add_note(
            note_data['student_id'],
            note_data['student_type'],
            note_data['classe'],
            note_data['matiere'],
            note_data['note']
        )
    
    return jsonify({'success': True})

@app.route('/vue_notes')
def vue_notes():
    notes = db.get_notes()
    classes = list(set([n['classe'] for n in notes]))
    matieres = list(set([n['matiere'] for n in notes]))
    
    return render_template('vue_notes.html', classes=classes, matieres=matieres)

@app.route('/get_notes_by_class', methods=['POST'])
def get_notes_by_class():
    data = request.json
    classe = data['classe']
    matiere = data['matiere']
    
    students = []
    
    if classe in ['maternelle', 'CI', 'CP', 'CE1', 'CE2', 'CM1', 'CM2']:
        all_students = db.get_ecoliers()
    else:
        all_students = db.get_eleves()
    
    for student in all_students:
        if student['classe'] == classe:
            student_notes = db.get_student_notes(student['id'], 'ecolier' if classe in ['maternelle', 'CI', 'CP', 'CE1', 'CE2', 'CM1', 'CM2'] else 'eleve')
            note_for_matiere = None
            for n in student_notes:
                if n['matiere'] == matiere:
                    note_for_matiere = n['note']
                    break
            
            students.append({
                'id': student['id'],
                'nom': student['nom'],
                'prenoms': student['prenoms'],
                'note': note_for_matiere
            })
    
    return jsonify({'students': students})

@app.route('/export_excel')
def export_excel():
    wb = openpyxl.Workbook()
    
    # Feuille pour écoliers
    ws_ecoliers = wb.active
    ws_ecoliers.title = "Écoliers"
    
    headers = ['ID', 'Nom', 'Prénoms', 'Sexe', 'Date de naissance', 'Classe', 
               'Numéro parents', 'Montant scolarité', 'Total payé', 'Reste', 'Date inscription']
    
    for col, header in enumerate(headers, 1):
        cell = ws_ecoliers.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    ecoliers = db.get_ecoliers()
    for row, ecolier in enumerate(ecoliers, 2):
        total_paid = db.get_total_paid(ecolier)
        reste = int(ecolier['montant_scolarite']) - total_paid
        
        ws_ecoliers.cell(row=row, column=1).value = ecolier['id']
        ws_ecoliers.cell(row=row, column=2).value = ecolier['nom']
        ws_ecoliers.cell(row=row, column=3).value = ecolier['prenoms']
        ws_ecoliers.cell(row=row, column=4).value = ecolier['sexe']
        ws_ecoliers.cell(row=row, column=5).value = ecolier['date_naissance']
        ws_ecoliers.cell(row=row, column=6).value = ecolier['classe']
        ws_ecoliers.cell(row=row, column=7).value = ecolier['numero_parents']
        ws_ecoliers.cell(row=row, column=8).value = ecolier['montant_scolarite']
        ws_ecoliers.cell(row=row, column=9).value = total_paid
        ws_ecoliers.cell(row=row, column=10).value = reste
        ws_ecoliers.cell(row=row, column=11).value = ecolier['date_inscription']
    
    # Feuille pour élèves
    ws_eleves = wb.create_sheet("Élèves")
    
    for col, header in enumerate(headers, 1):
        cell = ws_eleves.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    eleves = db.get_eleves()
    for row, eleve in enumerate(eleves, 2):
        total_paid = db.get_total_paid(eleve)
        reste = int(eleve['montant_scolarite']) - total_paid
        
        ws_eleves.cell(row=row, column=1).value = eleve['id']
        ws_eleves.cell(row=row, column=2).value = eleve['nom']
        ws_eleves.cell(row=row, column=3).value = eleve['prenoms']
        ws_eleves.cell(row=row, column=4).value = eleve['sexe']
        ws_eleves.cell(row=row, column=5).value = eleve['date_naissance']
        ws_eleves.cell(row=row, column=6).value = eleve['classe']
        ws_eleves.cell(row=row, column=7).value = eleve['numero_parents']
        ws_eleves.cell(row=row, column=8).value = eleve['montant_scolarite']
        ws_eleves.cell(row=row, column=9).value = total_paid
        ws_eleves.cell(row=row, column=10).value = reste
        ws_eleves.cell(row=row, column=11).value = eleve['date_inscription']
    
    # Feuille pour notes
    ws_notes = wb.create_sheet("Notes")
    note_headers = ['Étudiant', 'Classe', 'Matière', 'Note', 'Date']
    
    for col, header in enumerate(note_headers, 1):
        cell = ws_notes.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    all_notes = db.get_notes()
    for row, note in enumerate(all_notes, 2):
        student_name = ""
        if note['student_type'] == 'ecolier':
            for ecolier in ecoliers:
                if ecolier['id'] == note['student_id']:
                    student_name = f"{ecolier['nom']} {ecolier['prenoms']}"
                    break
        else:
            for eleve in eleves:
                if eleve['id'] == note['student_id']:
                    student_name = f"{eleve['nom']} {eleve['prenoms']}"
                    break
        
        ws_notes.cell(row=row, column=1).value = student_name
        ws_notes.cell(row=row, column=2).value = note['classe']
        ws_notes.cell(row=row, column=3).value = note['matiere']
        ws_notes.cell(row=row, column=4).value = note['note']
        ws_notes.cell(row=row, column=5).value = note['date']
    
    # Sauvegarder en mémoire
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        io.BytesIO(output.getvalue()),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'ecole_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=True)
  
