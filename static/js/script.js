function showSuccess(message) {
    const div = document.createElement('div');
    div.className = 'success-message';
    div.textContent = message;
    document.body.insertBefore(div, document.body.firstChild);
    
    setTimeout(() => {
        div.remove();
    }, 3000);
}

// Validation du formulaire
function validateForm(formData) {
    if (!formData.nom || !formData.prenoms || !formData.date_naissance || 
        !formData.numero_parents || !formData.montant_scolarite) {
        alert('Tous les champs sont obligatoires');
        return false;
    }
    
    if (!/^\d{8}$/.test(formData.numero_parents)) {
        alert('Le numéro doit contenir 8 chiffres');
        return false;
    }
    
    if (!/^\d+$/.test(formData.montant_scolarite)) {
        alert('Le montant doit être un nombre');
        return false;
    }
    
    return true;
}

// Fonction pour les paiements
function makePayment(studentId, studentType) {
    const amount = prompt('Entrez le montant à payer:');
    if (amount && /^\d+$/.test(amount)) {
        fetch('/paiement', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                student_id: studentId,
                student_type: studentType,
                amount: amount
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showSuccess('Paiement enregistré avec succès');
                setTimeout(() => location.reload(), 1000);
            }
        });
    }
}

// Fonctions pour les notes
let currentStudents = [];

function loadStudents() {
    const classe = document.getElementById('classe').value;
    const isEcolier = document.getElementById('type_student').value === 'ecolier';
    
    fetch('/get_students_by_class', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            classe: classe,
            is_ecolier: isEcolier
        })
    })
    .then(response => response.json())
    .then(data => {
        currentStudents = data.students;
        displayStudents();
    });
}

function displayStudents() {
    const container = document.getElementById('students-container');
    const matiere = document.getElementById('matiere').value;
    
    container.innerHTML = '';
    
    if (currentStudents.length === 0) {
        container.innerHTML = '<p>Aucun étudiant dans cette classe</p>';
        return;
    }
    
    const table = document.createElement('table');
    table.innerHTML = `
        <thead>
            <tr>
                <th>Nom</th>
                <th>Prénoms</th>
                <th>Note (${matiere})</th>
            </tr>
        </thead>
        <tbody id="students-tbody"></tbody>
    `;
    
    container.appendChild(table);
    
    const tbody = document.getElementById('students-tbody');
    currentStudents.forEach(student => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${student.nom}</td>
            <td>${student.prenoms}</td>
            <td>
                <input type="number" min="0" max="20" step="0.5" 
                       class="note-input" data-student-id="${student.id}"
                       data-student-type="${student.type}">
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function saveNotes() {
    const classe = document.getElementById('classe').value;
    const matiere = document.getElementById('matiere').value;
    const isEcolier = document.getElementById('type_student').value === 'ecolier';
    
    const notes = [];
    document.querySelectorAll('.note-input').forEach(input => {
        if (input.value) {
            notes.push({
                student_id: parseInt(input.dataset.studentId),
                student_type: input.dataset.studentType,
                classe: classe,
                matiere: matiere,
                note: input.value
            });
        }
    });
    
    if (notes.length === 0) {
        alert('Aucune note à enregistrer');
        return;
    }
    
    fetch('/save_notes', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ notes: notes })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            showSuccess('Notes enregistrées avec succès');
            setTimeout(() => location.reload(), 1000);
        }
    });
}

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    // Ajouter les écouteurs d'événements pour les formulaires
    const inscriptionForm = document.getElementById('inscription-form');
    if (inscriptionForm) {
        inscriptionForm.addEventListener('submit', function(e) {
            e.preventDefault();
            
            const formData = {
                nom: document.getElementById('nom').value,
                prenoms: document.getElementById('prenoms').value,
                sexe: document.getElementById('sexe').value,
                date_naissance: document.getElementById('date_naissance').value,
                classe: document.getElementById('classe').value,
                numero_parents: document.getElementById('numero_parents').value,
                montant_scolarite: document.getElementById('montant_scolarite').value,
                nom_enregistreur: document.getElementById('nom_enregistreur').value
            };
            
            if (!validateForm(formData)) return;
            
            const isEcolier = ['maternelle', 'CI', 'CP', 'CE1', 'CE2', 'CM1', 'CM2'].includes(formData.classe);
            const endpoint = isEcolier ? '/inscrire_ecolier' : '/inscrire_eleve';
            
            fetch(endpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(formData)
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showSuccess('Enregistrement effectué avec succès');
                    setTimeout(() => window.location.href = '/', 1500);
                }
            });
        });
    }
});
      
