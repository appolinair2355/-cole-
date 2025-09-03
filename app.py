import sys
print("Python version:", sys.version)

from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file
from models import Database
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
import io

app = Flask(__name__)
db = Database()

MATIERES = ['Mathématiques', 'Communication écrite', 'Lecture', 'Anglais', 'SVT', 'Histoire-géographie', 'Espagnol', 'EPS']

# ... (autres routes inchangées)

@app.route('/import_excel', methods=['GET', 'POST'])
def import_excel():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file)
            # Import Écoliers
            if 'Écoliers' in wb.sheetnames:
                ws = wb['Écoliers']
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[1] and row[2]:
                        ecolier = {
                            'nom': row[1],
                            'prenoms': row[2],
                            'sexe': row[3],
                            'date_naissance': row[4],
                            'classe': row[5],
                            'numero_parents': str(row[6]),
                            'montant_scolarite': str(row[7]),
                            'nom_enregistreur': 'Import Excel'
                        }
                        db.add_ecolier(ecolier)

            # Import Élèves
            if 'Élèves' in wb.sheetnames:
                ws = wb['Élèves']
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[1] and row[2]:
                        eleve = {
                            'nom': row[1],
                            'prenoms': row[2],
                            'sexe': row[3],
                            'date_naissance': row[4],
                            'classe': row[5],
                            'numero_parents': str(row[6]),
                            'montant_scolarite': str(row[7]),
                            'nom_enregistreur': 'Import Excel'
                        }
                        db.add_eleve(eleve)

            # Import Notes
            if 'Notes' in wb.sheetnames:
                ws = wb['Notes']
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[0] and row[1] and row[2] and row[3]:
                        # Trouver student_id à partir du nom
                        all_students = db.get_all()
                        for student in all_students:
                            full_name = f"{student['nom']} {student['prenoms']}"
                            if full_name == row[0]:
                                db.add_note(
                                    student['id'],
                                    student['type'],
                                    row[1],
                                    row[2],
                                    str(row[3])
                                )
                                break

            return jsonify({'success': True})
    return render_template('import_excel.html')

@app.route('/export_excel')
def export_excel():
    wb = openpyxl.Workbook()

    # Écoliers
    ws = wb.active
    ws.title = "Écoliers"
    headers = ['ID', 'Nom', 'Prénoms', 'Sexe', 'Date de naissance', 'Classe', 
               'Numéro parents', 'Montant scolarité', 'Total payé', 'Reste', 'Date inscription']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    ecoliers = db.get_ecoliers()
    for row, e in enumerate(ecoliers, 2):
        total = db.get_total_paid(e)
        reste = int(e['montant_scolarite']) - total
        ws.cell(row=row, column=1).value = e['id']
        ws.cell(row=row, column=2).value = e['nom']
        ws.cell(row=row, column=3).value = e['prenoms']
        ws.cell(row=row, column=4).value = e['sexe']
        ws.cell(row=row, column=5).value = e['date_naissance']
        ws.cell(row=row, column=6).value = e['classe']
        ws.cell(row=row, column=7).value = e['numero_parents']
        ws.cell(row=row, column=8).value = e['montant_scolarite']
        ws.cell(row=row, column=9).value = total
        ws.cell(row=row, column=10).value = reste
        ws.cell(row=row, column=11).value = e['date_inscription']

    # Élèves
    ws = wb.create_sheet("Élèves")
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    eleves = db.get_eleves()
    for row, e in enumerate(eleves, 2):
        total = db.get_total_paid(e)
        reste = int(e['montant_scolarite']) - total
        ws.cell(row=row, column=1).value = e['id']
        ws.cell(row=row, column=2).value = e['nom']
        ws.cell(row=row, column=3).value = e['prenoms']
        ws.cell(row=row, column=4).value = e['sexe']
        ws.cell(row=row, column=5).value = e['date_naissance']
        ws.cell(row=row, column=6).value = e['classe']
        ws.cell(row=row, column=7).value = e['numero_parents']
        ws.cell(row=row, column=8).value = e['montant_scolarite']
        ws.cell(row=row, column=9).value = total
        ws.cell(row=row, column=10).value = reste
        ws.cell(row=row, column=11).value = e['date_inscription']

    # Notes
    ws = wb.create_sheet("Notes")
    note_headers = ['Étudiant', 'Classe', 'Matière', 'Note', 'Date']
    for col, header in enumerate(note_headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    all_notes = db.get_notes()
    all_students = db.get_all()
    name_map = {s['id']: f"{s['nom']} {s['prenoms']}" for s in all_students}
    for row, n in enumerate(all_notes, 2):
        ws.cell(row=row, column=1).value = name_map.get(n['student_id'], 'Inconnu')
        ws.cell(row=row, column=2).value = n['classe']
        ws.cell(row=row, column=3).value = n['matiere']
        ws.cell(row=row, column=4).value = n['note']
        ws.cell(row=row, column=5).value = n['date']

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        io.BytesIO(output.getvalue()),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'ecole_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )

# ... (autres routes inchangées)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)
