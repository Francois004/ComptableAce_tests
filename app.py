from flask import Flask, render_template, request, redirect, url_for, send_file
import pdfkit
import pandas as pd
import openpyxl
import io
import pdfplumber
import PyPDF2
import re
import os
import fitz  # PyMuPDF
import easyocr

app = Flask(__name__)

# Nom du fichier Excel
FILENAME = 'clients.xlsx'

# Vérifie et initialise le fichier Excel s'il n'existe pas encore
def init_excel():
    try:
        # Tente de lire le fichier
        pd.read_excel(FILENAME)
    except FileNotFoundError:
        # Crée un nouveau fichier avec les colonnes nécessaires
        df = pd.DataFrame(columns=['Societe', 'Siren', 'Nom', 'Prenom', 'Email', 'Rib'])
        df.to_excel(FILENAME, index=False)

# Fonction pour ajouter des données dans le fichier Excel
def add_to_excel(data):
    df = pd.read_excel(FILENAME)
    df = df._append(data, ignore_index=True)
    df.to_excel(FILENAME, index=False)

# Route de la page principale avec le formulaire
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Récupération des données du formulaire
        societe = request.form['societe']
        siren = request.form['siren']
        nom = request.form['nom']
        prenom = request.form['prenom']
        email = request.form['email']
        rib = request.form['rib']
        
        # Validation basique (vous pouvez ajouter plus de validations si nécessaire)
        if not (societe and siren and nom and prenom and email and rib):
            return "Tous les champs sont obligatoires.", 400
        
        # Ajout des données au fichier Excel
        data = {
            'Société': societe,
            'SIREN': siren,
            'Nom': nom,
            'Prenom': prenom,
            'Email': email,
            'RIB': rib
        }
        add_to_excel(data)
        return redirect(url_for('view'))
    return render_template('index.html')

# Route pour visualiser et rechercher dans le fichier Excel
@app.route('/view', methods=['GET', 'POST'])
def view():
    df = pd.read_excel(FILENAME)
    
    # Gestion de la recherche si une requête est envoyée
    search_query = request.form.get('search_query', '').lower()
    if search_query:
        df = df[df.apply(lambda row: search_query in str(row['Nom']).lower() or search_query in str(row['Société']).lower(), axis=1)]
    
    # Conversion du DataFrame en liste de dictionnaires pour faciliter le rendu dans Jinja2
    data = df.to_dict(orient='records')
    print(data)
    return render_template('view.html', data=data)

#charger les données
df = pd.read_excel(FILENAME)
data = df.to_dict(orient='records')

@app.route('/generate_pdf/<int:client_id>')
def generate_pdf(client_id):
    client = data[client_id]  # Obtenir les informations du client

    # Rendre le modèle HTML avec les données du client
    html = render_template("mandat.html", client=client)

    # Convertir HTML en PDF
    pdf_buffer = pdfkit.from_string(html, False)

    # Envoyer le fichier PDF au navigateur pour téléchargement
    return send_file(
        io.BytesIO(pdf_buffer),
        as_attachment=True,
        download_name=f"mandat_{client['Nom']}.pdf",
        mimetype='application/pdf'
    )






# Fonction pour extraire le texte d'un PDF avec PyMuPDF
def extract_text_from_pdf(pdf_path):
    text = ""
    # Ouvrir le fichier PDF
    doc = fitz.open(pdf_path)
    zoom = 4
    mat = fitz.Matrix(zoom, zoom)
    count = 0
# Count variable is to get the number of pages in the pdf
    for p in doc:
      count += 1
    for i in range(count):
      val = f"image_{i+1}.png"
      page = doc.load_page(i)
      pix = page.get_pixmap(matrix=mat)
      pix.save(val)
    doc.close()
    reader = easyocr.Reader(['fr'],verbose=False )
    result = reader.readtext("image_1.png", detail=0)
    
    
    return result

def extract_data_line_by_line(text):
    extracted_data = []
    current_date = None
    current_operation = []
    current_debit = None

    date_pattern = r'(\d{2}/\d{2}/\d{4}) \| (\d{2}/\d{2}/\d{4})'  # Date au format dd/mm/yyyy | dd/mm/yyyy
    debit_pattern = r'(\d+\,\d{2})'  # Débit sous forme x,xx
    
    # Itérer sur chaque ligne du texte
    for line in text:
        # Recherche de la date (début de la ligne)
        date_match = re.match(date_pattern, line)
        if date_match:
            # Si une nouvelle date est trouvée, enregistrer les données extraites précédemment
            if current_date:
                extracted_data.append({
                    'Date de début': current_date[0],
                    'Date de fin': current_date[1],
                    'Opération': ' '.join(current_operation).strip(),
                    'Débit': current_debit
                })
            
            # Mettre à jour la date actuelle et réinitialiser l'opération et le débit
            current_date = (date_match.group(1), date_match.group(2))
            current_operation = []  # Réinitialiser l'opération pour la nouvelle ligne
            current_debit = None  # Réinitialiser le débit
            line = line[date_match.end():].strip()  # Supprimer la date de la ligne restante
        
        # Recherche du débit (s'il existe) dans la ligne
        debit_match = re.search(debit_pattern, line)
        if debit_match:
            current_debit = debit_match.group(1)
            line = line[:debit_match.start()].strip()  # Supprimer le débit de la ligne restante
        
        # Ajouter le reste du texte à l'opération
        if line:  # Assurez-vous qu'il y a encore du texte pour l'opération
            current_operation.append(line)
    
    # Ajouter la dernière ligne (si présente)
    if current_date:
        extracted_data.append({
            'Date de début': current_date[0],
            'Date de fin': current_date[1],
            'Opération': ' '.join(current_operation).strip(),
            'Débit': current_debit
        })

    return extracted_data


# Route pour uploader et traiter le fichier PDF
@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return "Fichier manquant", 400
    
    file = request.files['file']
    if file.filename == '':
        return "Fichier manquant", 400
    
    if file and file.filename.endswith('.pdf'):
        # Enregistrer le fichier PDF temporairement
        pdf_path = f"temp_{file.filename}"
        file.save(pdf_path)

        # Extraire le texte du fichier PDF
        full_text = extract_text_from_pdf(pdf_path)
       # print(full_text)
        
        # Vérifier si le texte a été extrait
        if not full_text:
            return "Impossible d'extraire du texte à partir du PDF", 400

        # Extraire les dates, opérations et débits
        data=extract_data_line_by_line(full_text)
        
        df = pd.DataFrame(data)

        # Sauvegarder en Excel
        excel_path = "output.xlsx"
        df.to_excel(excel_path, index=False)

        # Retourner le fichier Excel généré
        return send_file(excel_path, as_attachment=True, download_name="data.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return "Fichier invalide", 400


if __name__ == '__main__':
    init_excel()
    app.run(debug=True)
