import os
import sys
import logging
from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for
from models import Database
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
import io

logging.basicConfig(level=logging.INFO)
print("🚀 Démarrage de l’application École Mont Sion...")

app = Flask(__name__)
app.secret_key = 'ecole_mont_sion_secret_key'

db = Database()

MATIERES = [
    'Mathématiques',
    'Communication écrite',
    'Lecture',
    'Anglais',
    'SVT',
    'Histoire-géographie',
    'Espagnol',
    'EPS',
    'Conduite'
]

# ---------- ACCUEIL ----------
@app.route('/')
def accueil():
    return render_template('accueil.html')

# ---------- INSCRIPTION ----------
@app.route('/inscription')
def inscription():
    return render_template('inscription.html')

@app.route('/inscrire_ecolier', methods=['POST'])
def inscrire_ecolier():
    data = request.json
    db.add_ecolier({
        'nom': data['nom'],
        'prenoms': data['prenoms'],
        'sexe': data['sexe'],
        'date_naissance': data['date_naissance'],
        'classe': data['classe'],
        'numero_parents': data['numero_parents'],
        'montant_scolarite': data['montant_scolarite'],
        'nom_enregistreur': data['nom_enregistreur']
    })
    return jsonify({'success': True})

@app.route('/inscrire_eleve', methods=['POST'])
def inscrire_eleve():
    data = request.json
    db.add_eleve({
        'nom': data['nom'],
        'prenoms': data['prenoms'],
        'sexe': data['sexe'],
        'date_naissance': data['date_naissance'],
        'classe': data['classe'],
        'numero_parents': data['numero_parents'],
        'montant_scolarite': data['montant_scolarite'],
        'nom_enregistreur': data['nom_enregistreur']
    })
    return jsonify({'success': True})

# ---------- LISTES ----------
@app.route('/liste_eleves')
def liste_eleves():
    eleves = db.get_eleves()
    return render_template('liste_eleves.html', eleves=eleves)

@app.route('/liste_ecoliers')
def liste_ecoliers():
    ecoliers = db.get_ecoliers()
    return render_template('liste_ecoliers.html', ecoliers=ecoliers)

# ---------- SCOLARITÉ (Paiement automatique) ----------
@app.route('/scolarite')
def scolarite():
    students = db.get_all()
    for s in students:
        try:
            montant = int(str(s.get('montant_scolarite', '0')).strip())
        except ValueError:
            montant = 0
        total = db.get_total_paid(s)
        s['total_paid'] = total
        s['reste'] = montant - total
    return render_template('scolarite.html', students=students)

@app.route('/paiement', methods=['POST'])
def paiement():
    data = request.json
    success = db.add_payment(data['student_id'], data['student_type'], data['amount'])
    if success:
        # Recalculer le reste
        student = None
        if data['student_type'] == 'ecolier':
            for s in db.get_ecoliers():
                if s['id'] == data['student_id']:
                    student = s
                    break
        else:
            for s in db.get_eleves():
                if s['id'] == data['student_id']:
                    student = s
                    break
        
        if student:
            try:
                montant = int(str(student.get('montant_scolarite', '0')).strip())
            except ValueError:
                montant = 0
            total = db.get_total_paid(student)
            reste = montant - total
            return jsonify({'success': True, 'total_paid': total, 'reste': reste})
    
    return jsonify({'success': False})
