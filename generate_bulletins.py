#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generateur de Bulletins Scolaires
Lit notes.xlsx, applique la formule de reajustement, genere des bulletins HTML/PDF
"""

import os
from pathlib import Path
from openpyxl import load_workbook
from weasyprint import HTML
import unicodedata
import re
import random

# === CONFIGURATION ===

# Informations de l'etablissement
ETABLISSEMENT = {
    "nom": "Diploma Santé",
    "adresse": "85 Avenue Ledru Rollin",
    "code_postal": "75012",
    "ville": "Paris",
    "annee_scolaire": "2025/2026",
    "classe": "PAES",
    "charge_etudes": "Shirel Benchetrit",
    "semestre": "1er Semestre"
}

# Mapping des colonnes Excel vers les matieres du bulletin
# L'ordre correspond a l'ordre dans le template (MATIERE_1 a MATIERE_9)
MATIERES_CONFIG = [
    {"nom": "Biochimie", "col_excel": "Biochimie", "enseignant": "M. Da Fonseca"},
    {"nom": "Biologie Cellulaire", "col_excel": "Biologie Cellulaire", "enseignant": "M. Descatoire"},
    {"nom": "Biostatistiques", "col_excel": "Biostatistiques", "enseignant": "U. Bederede"},
    {"nom": "Chimie Médecine", "col_excel": "Chimie 1", "enseignant": "R. Hadjerci"},
    {"nom": "Chimie Terminale", "col_excel": "Chimie 2", "enseignant": "D. Yazidi"},
    {"nom": "Mathématiques", "col_excel": "Maths", "enseignant": "U. Bederede"},
    {"nom": "Physique", "col_excel": "Physique", "enseignant": "H. Diaw"},
    {"nom": "Physique/Biophysique", "col_excel": "Physique Biophysique", "enseignant": "H. Diaw"},
    {"nom": "SVT", "col_excel": "SVT", "enseignant": "M. Descatoire"},
]

# Appreciations par matiere - chaque matiere a ses propres appreciations
# pour eviter les doublons sur un meme bulletin
APPRECIATIONS_PAR_MATIERE = {
    "Biochimie": {
        "insuffisant": "Les notions de biochimie nécessitent un approfondissement. Un travail régulier permettra de progresser.",
        "passable": "Les bases en biochimie sont acquises. Poursuivez vos efforts pour consolider vos connaissances.",
        "assez_bien": "Bonne compréhension des concepts biochimiques. Continuez sur cette voie prometteuse.",
        "bien": "Très bonne maîtrise de la biochimie. Vos efforts sont récompensés, persévérez.",
        "excellent": "Excellente maîtrise des notions biochimiques. Félicitations pour ce travail remarquable."
    },
    "Biologie Cellulaire": {
        "insuffisant": "Les mécanismes cellulaires doivent être revus. Un travail plus soutenu est nécessaire.",
        "passable": "Compréhension correcte de la biologie cellulaire. Des efforts supplémentaires consolideront vos acquis.",
        "assez_bien": "Bonne assimilation des concepts cellulaires. Maintenez cette dynamique positive.",
        "bien": "Très bonne compréhension des processus cellulaires. Continuez ainsi.",
        "excellent": "Maîtrise remarquable de la biologie cellulaire. Travail exemplaire."
    },
    "Biostatistiques": {
        "insuffisant": "Les méthodes statistiques demandent plus de pratique. Un entraînement régulier est conseillé.",
        "passable": "Les bases statistiques sont comprises. Continuez à vous exercer pour gagner en aisance.",
        "assez_bien": "Bonne application des outils statistiques. Poursuivez vos efforts.",
        "bien": "Très bonne maîtrise des biostatistiques. Résultats très satisfaisants.",
        "excellent": "Excellente compréhension et application des méthodes statistiques. Bravo."
    },
    "Chimie Médecine": {
        "insuffisant": "Les fondamentaux en chimie médicale doivent être renforcés. Travaillez régulièrement.",
        "passable": "Niveau correct en chimie médicale. Poursuivez vos efforts pour vous améliorer.",
        "assez_bien": "Bonne compréhension de la chimie appliquée à la médecine. Continuez ainsi.",
        "bien": "Très bon niveau en chimie médicale. Vos résultats sont encourageants.",
        "excellent": "Excellente maîtrise de la chimie médicale. Félicitations pour votre investissement."
    },
    "Chimie Terminale": {
        "insuffisant": "Les acquis de chimie terminale nécessitent une révision. Un travail soutenu s'impose.",
        "passable": "Les notions de chimie terminale sont assimilées. Continuez à progresser.",
        "assez_bien": "Bonne maîtrise des concepts de chimie terminale. Persévérez dans vos efforts.",
        "bien": "Très bonne compréhension de la chimie terminale. Résultats très positifs.",
        "excellent": "Excellents résultats en chimie terminale. Travail remarquable et rigoureux."
    },
    "Mathématiques": {
        "insuffisant": "Les compétences mathématiques doivent être consolidées. Un entraînement quotidien est recommandé.",
        "passable": "Niveau satisfaisant en mathématiques. Continuez à pratiquer pour progresser.",
        "assez_bien": "Bonne maîtrise des outils mathématiques. Maintenez vos efforts.",
        "bien": "Très bon niveau en mathématiques. Vos compétences sont solides.",
        "excellent": "Excellente maîtrise des mathématiques. Résultats impressionnants, félicitations."
    },
    "Physique": {
        "insuffisant": "Les concepts physiques nécessitent plus de travail. Revoyez les notions fondamentales.",
        "passable": "Compréhension correcte des phénomènes physiques. Poursuivez vos efforts.",
        "assez_bien": "Bonne assimilation des lois physiques. Continuez sur cette lancée positive.",
        "bien": "Très bonne maîtrise de la physique. Vos efforts portent leurs fruits.",
        "excellent": "Excellente compréhension de la physique. Travail exemplaire et rigoureux."
    },
    "Physique/Biophysique": {
        "insuffisant": "Les notions de biophysique demandent un approfondissement. Travaillez régulièrement.",
        "passable": "Les bases en biophysique sont acquises. Continuez à consolider vos connaissances.",
        "assez_bien": "Bonne compréhension des applications physiques en biologie. Persévérez.",
        "bien": "Très bon niveau en biophysique. Résultats très encourageants.",
        "excellent": "Maîtrise excellente de la biophysique. Félicitations pour ce parcours remarquable."
    },
    "SVT": {
        "insuffisant": "Les connaissances en SVT doivent être renforcées. Un travail plus régulier est nécessaire.",
        "passable": "Niveau correct en SVT. Poursuivez vos efforts pour améliorer vos résultats.",
        "assez_bien": "Bonne compréhension des sciences de la vie et de la Terre. Continuez ainsi.",
        "bien": "Très bonne maîtrise des SVT. Vos résultats reflètent un travail sérieux.",
        "excellent": "Excellents résultats en SVT. Travail remarquable et approfondi, bravo."
    }
}

# Chemins
BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "notes.xlsx"
TEMPLATE_FILE = BASE_DIR / "bulletin_template.html"
HTML_OUTPUT_DIR = BASE_DIR / "bulletins_html"
PDF_OUTPUT_DIR = BASE_DIR / "bulletins_pdf"

def adjust_grade(note):
    """Applique la formule f(x) = 0.57x + 8.74, plafonnee a 20
    Si pas de note, retourne une note aleatoire entre 9.5 et 12"""
    if note is None or note == "-" or note == "":
        # Generer une note aleatoire entre 9.5 et 12 (note APRES harmonisation)
        return round(random.uniform(9.5, 12.0), 1)
    try:
        note_float = float(note)
        adjusted = 0.57 * note_float + 8.74
        return min(adjusted, 20.0)  # Plafonner a 20
    except (ValueError, TypeError):
        # Si erreur de conversion, generer une note aleatoire
        return round(random.uniform(9.5, 12.0), 1)


def get_appreciation(note, matiere_nom):
    """Genere une appreciation basee sur la note et la matiere"""
    if note is None:
        note = 10.5  # Note moyenne par defaut

    # Recuperer les appreciations de la matiere
    appreciations = APPRECIATIONS_PAR_MATIERE.get(matiere_nom, {})

    if note < 10:
        return appreciations.get("insuffisant", "Des efforts sont nécessaires pour progresser.")
    elif note < 12:
        return appreciations.get("passable", "Résultats encourageants. Continuez vos efforts.")
    elif note < 14:
        return appreciations.get("assez_bien", "Bon travail. Poursuivez dans cette voie.")
    elif note < 16:
        return appreciations.get("bien", "Très bon travail. Continuez ainsi.")
    else:
        return appreciations.get("excellent", "Excellents résultats. Félicitations.")


def get_appreciation_generale(prenom, moyenne):
    """Genere une appreciation generale personnalisee"""
    if moyenne is None:
        moyenne = 10.5

    if moyenne < 10:
        return f"{prenom} doit fournir davantage d'efforts pour progresser. Un travail régulier et soutenu permettra d'améliorer les résultats au prochain semestre."
    elif moyenne < 12:
        return f"{prenom} montre des résultats encourageants. Avec plus de régularité dans le travail, les résultats continueront de s'améliorer."
    elif moyenne < 14:
        return f"{prenom} fournit un bon travail ce semestre. Les bases sont acquises et les efforts doivent être maintenus pour progresser davantage."
    elif moyenne < 16:
        return f"{prenom} a fourni un travail régulier et rigoureux tout au long du semestre. Les résultats sont très satisfaisants. Continuez ainsi."
    else:
        return f"{prenom} a fourni un travail remarquable tout au long du semestre. Félicitations pour ces excellents résultats."


def format_note(note):
    """Formate une note pour l'affichage"""
    if note is None:
        return "NN"
    return f"{note:.1f}".replace(".", ",")


def format_date(date_value):
    """Formate une date pour l'affichage"""
    if date_value is None:
        return ""
    if hasattr(date_value, 'strftime'):
        return date_value.strftime("%d/%m/%Y")
    return str(date_value)


def sanitize_filename(text):
    """Nettoie un texte pour en faire un nom de fichier valide"""
    # Normaliser les caracteres unicode
    text = unicodedata.normalize('NFKD', str(text))
    text = text.encode('ASCII', 'ignore').decode('ASCII')
    # Remplacer les espaces et caracteres speciaux
    text = re.sub(r'[^\w\s-]', '', text)
    text = re.sub(r'[-\s]+', '_', text)
    return text.lower()


def load_excel_data():
    """Charge les donnees depuis le fichier Excel"""
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active

    # Lire les en-tetes
    headers = []
    for cell in ws[1]:
        headers.append(cell.value)

    # Lire les donnees des eleves
    students = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:  # Ligne vide
            continue

        student = {
            "nom": row[0],
            "prenom": row[1],
            "date_naissance": row[2],
            "email": row[3],
            "choix_parcoursup": row[4],
            "notes": {}
        }

        # Lire les notes (colonnes 5 a 13, index 0-based: 5-13)
        note_columns = ["Biochimie", "Biologie Cellulaire", "Biostatistiques",
                       "Chimie 1", "Chimie 2", "Maths", "Physique",
                       "Physique Biophysique", "SVT"]

        for i, col_name in enumerate(note_columns):
            col_index = 5 + i  # Les notes commencent a la colonne 6 (index 5)
            if col_index < len(row):
                student["notes"][col_name] = row[col_index]

        students.append(student)

    return students


def calculate_class_stats(students):
    """Calcule les statistiques de classe pour chaque matiere"""
    stats = {}

    for matiere in MATIERES_CONFIG:
        col_name = matiere["col_excel"]
        notes = []

        for student in students:
            raw_note = student["notes"].get(col_name)
            adjusted = adjust_grade(raw_note)
            if adjusted is not None:
                notes.append(adjusted)

        if notes:
            stats[col_name] = {
                "moyenne": sum(notes) / len(notes),
                "min": min(notes),
                "max": max(notes)
            }
        else:
            stats[col_name] = {
                "moyenne": None,
                "min": None,
                "max": None
            }

    return stats


def generate_bulletin_html(student, class_stats, template):
    """Genere le HTML d'un bulletin pour un eleve"""
    html = template

    # Informations etablissement
    html = html.replace("{{NOM_ETABLISSEMENT}}", ETABLISSEMENT["nom"])
    html = html.replace("{{ADRESSE_ETABLISSEMENT}}", ETABLISSEMENT["adresse"])
    html = html.replace("{{CODE_POSTAL}}", ETABLISSEMENT["code_postal"])
    html = html.replace("{{VILLE}}", ETABLISSEMENT["ville"])
    html = html.replace("{{ANNEE_SCOLAIRE}}", ETABLISSEMENT["annee_scolaire"])
    html = html.replace("{{CLASSE}}", ETABLISSEMENT["classe"])
    html = html.replace("{{CHARGE_ETUDES}}", ETABLISSEMENT["charge_etudes"])
    html = html.replace("{{SEMESTRE}}", ETABLISSEMENT["semestre"])

    # Informations eleve
    html = html.replace("{{PRENOM_ELEVE}}", str(student["prenom"] or ""))
    html = html.replace("{{NOM_ELEVE}}", str(student["nom"] or ""))
    html = html.replace("{{DATE_NAISSANCE}}", format_date(student["date_naissance"]))

    # Notes et appreciations par matiere
    student_notes = []

    for i, matiere in enumerate(MATIERES_CONFIG, start=1):
        col_name = matiere["col_excel"]
        raw_note = student["notes"].get(col_name)
        adjusted_note = adjust_grade(raw_note)

        html = html.replace(f"{{{{MATIERE_{i}}}}}", matiere["nom"])
        html = html.replace(f"{{{{ENSEIGNANT_{i}}}}}", matiere["enseignant"])
        html = html.replace(f"{{{{MOY_ELEVE_{i}}}}}", format_note(adjusted_note))

        # Statistiques de classe
        stats = class_stats.get(col_name, {})
        html = html.replace(f"{{{{MOY_CLASSE_{i}}}}}", format_note(stats.get("moyenne")))
        html = html.replace(f"{{{{NOTE_MIN_{i}}}}}", format_note(stats.get("min")))
        html = html.replace(f"{{{{NOTE_MAX_{i}}}}}", format_note(stats.get("max")))

        # Appreciation specifique a la matiere
        html = html.replace(f"{{{{APPRECIATION_{i}}}}}", get_appreciation(adjusted_note, matiere["nom"]))

        if adjusted_note is not None:
            student_notes.append(adjusted_note)

    # Moyennes generales
    if student_notes:
        moyenne_eleve = sum(student_notes) / len(student_notes)
    else:
        moyenne_eleve = 10.5

    # Calculer moyenne generale de la classe
    all_class_moyennes = []
    for matiere in MATIERES_CONFIG:
        col_name = matiere["col_excel"]
        stats = class_stats.get(col_name, {})
        if stats.get("moyenne") is not None:
            all_class_moyennes.append(stats["moyenne"])

    if all_class_moyennes:
        moyenne_classe = sum(all_class_moyennes) / len(all_class_moyennes)
    else:
        moyenne_classe = None

    html = html.replace("{{MOYENNE_GENERALE_ELEVE}}", format_note(moyenne_eleve))
    html = html.replace("{{MOYENNE_GENERALE_CLASSE}}", format_note(moyenne_classe))

    # Absences et appreciation generale
    html = html.replace("{{ABSENCES}}", "RAS")
    html = html.replace("{{APPRECIATION_GENERALE}}",
                       get_appreciation_generale(student["prenom"], moyenne_eleve))

    return html


def generate_index_html(students):
    """Genere la page index.html avec la liste des eleves"""

    # Trier les eleves par nom de famille
    sorted_students = sorted(students, key=lambda s: (s["nom"] or "", s["prenom"] or ""))

    rows = []
    for student in sorted_students:
        filename = f"bulletin_{sanitize_filename(student['prenom'])}_{sanitize_filename(student['nom'])}.pdf"
        pdf_path = f"bulletins_pdf/{filename}"

        row = f"""
            <tr>
                <td>{student['prenom'] or ''}</td>
                <td>{student['nom'] or ''}</td>
                <td><a href="mailto:{student['email'] or ''}">{student['email'] or ''}</a></td>
                <td class="pdf-cell">
                    <a href="{pdf_path}" class="btn-pdf" target="_blank" title="Voir le bulletin">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                            <path d="M14 14V4.5L9.5 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2zM9.5 3A1.5 1.5 0 0 0 11 4.5h2V14a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1h5.5v2z"/>
                            <path d="M4.603 14.087a.81.81 0 0 1-.438-.42c-.195-.388-.13-.776.08-1.102.198-.307.526-.568.897-.787a7.68 7.68 0 0 1 1.482-.645 19.697 19.697 0 0 0 1.062-2.227 7.269 7.269 0 0 1-.43-1.295c-.086-.4-.119-.796-.046-1.136.075-.354.274-.672.65-.823.192-.077.4-.12.602-.077a.7.7 0 0 1 .477.365c.088.164.12.356.127.538.007.188-.012.396-.047.614-.084.51-.27 1.134-.52 1.794a10.954 10.954 0 0 0 .98 1.686 5.753 5.753 0 0 1 1.334.05c.364.066.734.195.96.465.12.144.193.32.2.518.007.192-.047.382-.138.563a1.04 1.04 0 0 1-.354.416.856.856 0 0 1-.51.138c-.331-.014-.654-.196-.933-.417a5.712 5.712 0 0 1-.911-.95 11.651 11.651 0 0 0-1.997.406 11.307 11.307 0 0 1-1.02 1.51c-.292.35-.609.656-.927.787a.793.793 0 0 1-.58.029z"/>
                        </svg>
                        PDF
                    </a>
                    <a href="{pdf_path}" class="btn-download" download title="Télécharger">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                            <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/>
                            <path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z"/>
                        </svg>
                    </a>
                </td>
            </tr>"""
        rows.append(row)

    index_html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bulletins Scolaires - {ETABLISSEMENT['nom']}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Open Sans', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}

        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }}

        .header {{
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 30px 40px;
            text-align: center;
        }}

        .header h1 {{
            font-size: 28px;
            margin-bottom: 10px;
        }}

        .header p {{
            font-size: 14px;
            opacity: 0.9;
        }}

        .stats {{
            display: flex;
            justify-content: center;
            gap: 40px;
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.2);
        }}

        .stat {{
            text-align: center;
        }}

        .stat-value {{
            font-size: 32px;
            font-weight: 700;
        }}

        .stat-label {{
            font-size: 12px;
            opacity: 0.8;
        }}

        .search-container {{
            padding: 20px 40px;
            background: #f8f9fa;
            border-bottom: 1px solid #eee;
        }}

        .search-input {{
            width: 100%;
            padding: 12px 20px;
            font-size: 14px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            outline: none;
            transition: border-color 0.3s;
        }}

        .search-input:focus {{
            border-color: #3498db;
        }}

        .table-container {{
            padding: 20px 40px 40px;
            overflow-x: auto;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
        }}

        th {{
            background: #f8f9fa;
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            color: #2c3e50;
            border-bottom: 2px solid #dee2e6;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        td {{
            padding: 15px 12px;
            border-bottom: 1px solid #eee;
            vertical-align: middle;
        }}

        tr:hover {{
            background: #f8f9fa;
        }}

        a {{
            color: #3498db;
            text-decoration: none;
        }}

        a:hover {{
            text-decoration: underline;
        }}

        .pdf-cell {{
            display: flex;
            gap: 8px;
            align-items: center;
        }}

        .btn-pdf, .btn-download {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            text-decoration: none;
            transition: all 0.2s;
        }}

        .btn-pdf {{
            background: #e74c3c;
            color: white;
        }}

        .btn-pdf:hover {{
            background: #c0392b;
            text-decoration: none;
        }}

        .btn-download {{
            background: #27ae60;
            color: white;
            padding: 8px;
        }}

        .btn-download:hover {{
            background: #1e8449;
            text-decoration: none;
        }}

        .no-results {{
            text-align: center;
            padding: 40px;
            color: #666;
            display: none;
        }}

        @media (max-width: 768px) {{
            .container {{
                border-radius: 0;
            }}

            .header, .search-container, .table-container {{
                padding-left: 20px;
                padding-right: 20px;
            }}

            .stats {{
                flex-direction: column;
                gap: 15px;
            }}

            th, td {{
                padding: 10px 8px;
                font-size: 12px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Bulletins du {ETABLISSEMENT['semestre']}</h1>
            <p>{ETABLISSEMENT['nom']} - Année scolaire {ETABLISSEMENT['annee_scolaire']} - Classe {ETABLISSEMENT['classe']}</p>
            <div class="stats">
                <div class="stat">
                    <div class="stat-value">{len(students)}</div>
                    <div class="stat-label">Élèves</div>
                </div>
                <div class="stat">
                    <div class="stat-value">{len(MATIERES_CONFIG)}</div>
                    <div class="stat-label">Matières</div>
                </div>
            </div>
        </div>

        <div class="search-container">
            <input type="text" class="search-input" id="searchInput"
                   placeholder="Rechercher un élève par nom, prénom ou email...">
        </div>

        <div class="table-container">
            <table id="studentsTable">
                <thead>
                    <tr>
                        <th>Prénom</th>
                        <th>Nom</th>
                        <th>Email</th>
                        <th>Bulletin</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(rows)}
                </tbody>
            </table>
            <div class="no-results" id="noResults">
                Aucun élève trouvé pour cette recherche.
            </div>
        </div>
    </div>

    <script>
        document.getElementById('searchInput').addEventListener('input', function() {{
            const searchTerm = this.value.toLowerCase();
            const rows = document.querySelectorAll('#studentsTable tbody tr');
            let visibleCount = 0;

            rows.forEach(row => {{
                const text = row.textContent.toLowerCase();
                const isVisible = text.includes(searchTerm);
                row.style.display = isVisible ? '' : 'none';
                if (isVisible) visibleCount++;
            }});

            document.getElementById('noResults').style.display =
                visibleCount === 0 ? 'block' : 'none';
        }});
    </script>
</body>
</html>"""

    return index_html


def main():
    print("=" * 60)
    print("GENERATEUR DE BULLETINS SCOLAIRES")
    print("=" * 60)

    # Creer les dossiers de sortie
    HTML_OUTPUT_DIR.mkdir(exist_ok=True)
    PDF_OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"\n[OK] Dossiers de sortie créés")

    # Charger le template
    print(f"[...] Chargement du template...")
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()
    print(f"[OK] Template chargé: {TEMPLATE_FILE}")

    # Charger les donnees Excel
    print(f"[...] Chargement des données Excel...")
    students = load_excel_data()
    print(f"[OK] {len(students)} élèves chargés depuis {EXCEL_FILE}")

    # Calculer les statistiques de classe
    print(f"[...] Calcul des statistiques de classe...")
    class_stats = calculate_class_stats(students)
    print(f"[OK] Statistiques calculées pour {len(class_stats)} matières")

    # Generer les bulletins
    print(f"\n[...] Génération des bulletins...")

    for i, student in enumerate(students, start=1):
        prenom = student['prenom'] or 'inconnu'
        nom = student['nom'] or 'inconnu'
        filename_base = f"bulletin_{sanitize_filename(prenom)}_{sanitize_filename(nom)}"

        # Generer le HTML
        html_content = generate_bulletin_html(student, class_stats, template)
        html_path = HTML_OUTPUT_DIR / f"{filename_base}.html"

        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Convertir en PDF
        pdf_path = PDF_OUTPUT_DIR / f"{filename_base}.pdf"
        HTML(string=html_content, base_url=str(BASE_DIR)).write_pdf(pdf_path)

        print(f"  [{i}/{len(students)}] {prenom} {nom} - OK")

    print(f"\n[OK] {len(students)} bulletins générés")

    # Generer l'index HTML
    print(f"\n[...] Génération de la plateforme index.html...")
    index_html = generate_index_html(students)
    index_path = BASE_DIR / "index.html"

    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(index_html)

    print(f"[OK] Plateforme créée: {index_path}")

    # Resume
    print("\n" + "=" * 60)
    print("RÉSUMÉ")
    print("=" * 60)
    print(f"  Bulletins HTML: {HTML_OUTPUT_DIR}")
    print(f"  Bulletins PDF:  {PDF_OUTPUT_DIR}")
    print(f"  Plateforme:     {index_path}")
    print(f"\nOuvrez index.html dans votre navigateur pour accéder aux bulletins.")
    print("=" * 60)


if __name__ == "__main__":
    main()
