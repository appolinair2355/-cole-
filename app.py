import os, uuid, re, datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import yaml
from openpyxl import Workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'ecole-mont-sion-secret-key'

# CONSTANTES
PORT = 10000
DATABASE = 'database.yaml'
MATIERES = ['Math', 'Français', 'Science', 'Histoire', 'Géographie', 'Anglais', 'EPS']
TRIMESTRES = ['Intero1', 'Intero2']
ALLOWED_EXTENSIONS = {'xlsx'}

# ┌─────────────────────────────┐
# │ 1. GESTION BASE YAML        │
# └─────────────────────────────┘
def load_data():
    if os.path.exists(DATABASE):
        with open(DATABASE, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f) or {'primaire': [], 'secondaire': []}
    else:
        data = {'primaire': [], 'secondaire': []}
    for niveau in ['primaire', 'secondaire']:
        for eleve in data[niveau]:
            eleve.setdefault('notes', {})
            eleve.setdefault('paiements', [])
            for m in MATIERES:
                eleve['notes'].setdefault(m, {})
                for t in TRIMESTRES:
                    eleve['notes'][m].setdefault(t, '')
    return data

def save_data(data):
    with open(DATABASE, 'w', encoding='utf-8') as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)

# ┌─────────────────────────────┐
# │ 2. ROUTES WEB               │
# └─────────────────────────────┘
@app.route('/')
def index():
    return render_template('index.html')

# -----------------------------
# INSCRIPTION
# -----------------------------
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        nom = request.form.get('nom', '').strip()
        prenom = request.form.get('prenom', '').strip()
        classe = request.form.get('classe', '').strip()
        niveau = request.form.get('niveau', '').strip()  # 'primaire' ou 'secondaire'
        if not nom or not prenom or not classe or niveau not in ['primaire', 'secondaire']:
            flash('Tous les champs sont obligatoires.', 'danger')
            return redirect(url_for('register'))
        data = load_data()
        eleve = {
            'id': str(uuid.uuid4()),
            'nom': nom,
            'prenom': prenom,
            'classe': classe,
            'notes': {},
            'paiements': []
        }
        for m in MATIERES:
            eleve['notes'][m] = {t: '' for t in TRIMESTRES}
        data[niveau].append(eleve)
        save_data(data)
        flash('Élève inscrit avec succès.', 'success')
        return redirect(url_for('students'))
    return render_template('register.html')

# -----------------------------
# LISTE ÉLÈVES
# -----------------------------
@app.route('/students')
def students():
    data = load_data()
    classes = {}
    for niveau in ['primaire', 'secondaire']:
        for eleve in data[niveau]:
            classe = eleve['classe']
            classes.setdefault(classe, []).append(eleve)
    return render_template('students.html', classes=classes)

# -----------------------------
# MODIFIER / SUPPRIMER
# -----------------------------
@app.route('/edit/<id>', methods=['GET', 'POST'])
def edit(id):
    data = load_data()
    eleve = None
    niveau = ''
    for n in ['primaire', 'secondaire']:
        for e in data[n]:
            if e['id'] == id:
                eleve = e
                niveau = n
                break
    if not eleve:
        flash('Élève introuvable.', 'danger')
        return redirect(url_for('students'))
    if request.method == 'POST':
        eleve['nom'] = request.form.get('nom', '').strip()
        eleve['prenom'] = request.form.get('prenom', '').strip()
        eleve['classe'] = request.form.get('classe', '').strip()
        save_data(data)
        flash('Élève modifié avec succès.', 'success')
        return redirect(url_for('students'))
    return render_template('edit.html', eleve=eleve)

@app.route('/delete/<id>', methods=['POST'])
def delete(id):
    data = load_data()
    for n in ['primaire', 'secondaire']:
        data[n] = [e for e in data[n] if e['id'] != id]
    save_data(data)
    flash('Élève supprimé avec succès.', 'success')
    return redirect(url_for('students'))

# -----------------------------
# SCOLARITÉ
# -----------------------------
@app.route('/scolarite')
def scolarite():
    data = load_data()
    eleves = []
    for n in ['primaire', 'secondaire']:
        for e in data[n]:
            paye = sum(p['montant'] for p in e['paiements'])
            reste = 1000 - paye  # frais annuel fixe
            eleves.append({'eleve': e, 'paye': paye, 'reste': reste})
    return render_template('scolarite.html', eleves=eleves)

@app.route('/scolarite/payer', methods=['POST'])
def payer():
    eleve_id = request.form.get('eleve_id')
    montant = float(request.form.get('montant', 0))
    data = load_data()
    for n in ['primaire', 'secondaire']:
        for e in data[n]:
            if e['id'] == eleve_id:
                e['paiements'].append({
                    'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                    'montant': montant
                })
                save_data(data)
                flash('Paiement enregistré.', 'success')
                return redirect(url_for('scolarite'))
    flash('Élève introuvable.', 'danger')
    return redirect(url_for('scolarite'))

# -----------------------------
# NOTES
# -----------------------------
@app.route('/notes', methods=['GET', 'POST'])
def notes():
    data = load_data()
    if request.method == 'POST':
        for n in ['primaire', 'secondaire']:
            for e in data[n]:
                for m in MATIERES:
                    for t in TRIMESTRES:
                        val = request.form.get(f"{e['id']}_{m}_{t}", '').strip()
                        e['notes'][m][t] = val
        save_data(data)
        flash('Notes enregistrées avec succès.', 'success')
        return redirect(url_for('notes'))
    eleves = []
    for n in ['primaire', 'secondaire']:
        eleves.extend(data[n])
    return render_template('notes.html', eleves=eleves, matieres=MATIERES, trimestres=TRIMESTRES)

# -----------------------------
# IMPORT EXCEL (fusionne sans écraser)
# -----------------------------
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/import_excel', methods=['GET', 'POST'])
def import_excel():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            flash('Aucun fichier sélectionné.', 'danger')
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash('Fichier non autorisé ( uniquement .xlsx ).', 'danger')
            return redirect(request.url)
        try:
            from openpyxl import load_workbook
            wb = load_workbook(file)
            ws = wb.active
            data = load_data()
            # On suppose :
            # A:nom, B:prenom, C:classe, D:niveau (primaire/secondaire), E:paiement, F+ : notes
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0] or not row[1] or not row[2] or not row[3]:
                    continue
                nom, prenom, classe, niveau = str(row[0]).strip(), str(row[1]).strip(), str(row[2]).strip(), str(row[3]).strip().lower()
                if niveau not in ['primaire', 'secondaire']:
                    continue
                paye = float(row[4]) if row[4] else 0
                eleve = {
                    'id': str(uuid.uuid4()),
                    'nom': nom,
                    'prenom': prenom,
                    'classe': classe,
                    'notes': {},
                    'paiements': []
                }
                if paye:
                    eleve['paiements'].append({'date': datetime.datetime.now().strftime('%Y-%m-%d'), 'montant': paye})
                for m in MATIERES:
                    eleve['notes'][m] = {}
                    for t in TRIMESTRES:
                        eleve['notes'][m][t] = ''
                # notes
                idx = 5
                for m in MATIERES:
                    for t in TRIMESTRES:
                        if idx < len(row) and row[idx]:
                            eleve['notes'][m][t] = str(row[idx]).strip()
                        idx += 1
                data[niveau].append(eleve)
            save_data(data)
            flash('Importation réussie avec succès.', 'success')
            return redirect(url_for('students'))
        except Exception as e:
            flash(f'Erreur lors de l\'import : {e}', 'danger')
            return redirect(url_for('import_excel'))
    return render_template('import_excel.html')

# -----------------------------
# EXPORT EXCEL – TOUTE LA BASE
# -----------------------------
@app.route('/export_excel')
def export_excel():
    try:
        data = load_data()
        wb = Workbook()
        wb.remove(wb.active)  # enlève feuille vide
        # Regroupement par classe
        classes = {}
        for n in ['primaire', 'secondaire']:
            for e in data[n]:
                classe = e['classe']
                classes.setdefault(classe, []).append(e)
        for classe, eleves in classes.items():
            ws = wb.create_sheet(title=classe[:31])  # max 31 car
            # En-têtes
            headers = ['Nom', 'Prénom', 'Classe']
            for m in MATIERES:
                for t in TRIMESTRES:
                    headers.append(f"{m}_{t}")
            headers += ['Total Payé', 'Reste']
            ws.append(headers)
            # Lignes
            for e in eleves:
                row = [e['nom'], e['prenom'], e['classe']]
                for m in MATIERES:
                    for t in TRIMESTRES:
                        row.append(e['notes'][m][t])
                paye = sum(p['montant'] for p in e['paiements'])
                reste = 1000 - paye
                row += [paye, reste]
                ws.append(row)
        # Fichier temp
        from tempfile import NamedTemporaryFile
        with NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            wb.save(tmp.name)
            tmp_path = tmp.name
        flash('Exportation réussie avec succès.', 'success')
        return send_file(tmp_path, as_attachment=True, download_name='ecole_mont_sion_complet.xlsx')
    except Exception as e:
        flash(f'Erreur lors de l\'export : {e}', 'danger')
        return redirect(url_for('students'))

# -----------------------------
# LANCEMENT
# -----------------------------
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=PORT, debug=True)
