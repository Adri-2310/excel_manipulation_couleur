from flask import Flask, render_template, request, jsonify, after_this_request
import pandas as pd
import os
import glob
import traceback
from flask import send_from_directory

# <-- importe ton script de logique ici (adapter le nom du fichier / fonction) -->
from logic import apply_colors_to_file2

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    try:
        if request.method == 'POST':
            # ----- FICHIER SOURCE (OBLIGATOIRE) -----
            if 'file' not in request.files:
                return "Aucun fichier source téléversé.", 400

            file_source = request.files['file']
            if file_source.filename == '':
                return "Nom de fichier source vide.", 400

            source_path = os.path.join(app.config['UPLOAD_FOLDER'], file_source.filename)
            file_source.save(source_path)

            sheet_source = request.form.get('sheet_name')
            if not sheet_source:
                return "Aucune feuille source sélectionnée.", 400

            # ----- FICHIER CIBLE (OBLIGATOIRE OU NON, À TOI DE VOIR) -----
            file_target = request.files.get('file_target')
            if not file_target or file_target.filename == '':
                return "Aucun fichier cible téléversé.", 400

            target_path = os.path.join(app.config['UPLOAD_FOLDER'], file_target.filename)
            file_target.save(target_path)

            sheet_target = request.form.get('sheet_name_target')
            if not sheet_target:
                return "Aucune feuille cible sélectionnée.", 400
            # ----- APPEL DE TON SCRIPT PYTHON DE LOGIQUE -----
            # Ici tu appelles la fonction qui fait le vrai travail sur les Excels
            # À adapter en fonction de la signature réelle de ta fonction
            apply_colors_to_file2(file1_path=file_source,file1_sheet=sheet_source,file2_path=file_target,file2_sheet=sheet_target)

            # Option 2 (alternative) : afficher un message / une page
            message = f"Traitement terminé."
            return render_template(
                'resultat.html',
                message=message,
                chemin_fichier_modifie=os.path.basename(target_path),
            )

        return render_template('index.html', table=None, filename=None, sheet_name=None)
    except Exception as e:
        print(f"Erreur dans index: {traceback.format_exc()}")
        return f"Une erreur est survenue : {str(e)}", 500

@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Aucun fichier téléversé'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Nom de fichier vide'}), 400

        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        return jsonify({'sheet_names': sheet_names})
    except Exception as e:
        print(f"Erreur dans get_sheets: {traceback.format_exc()}")
        return jsonify({'error': f'Erreur lors de la lecture du fichier : {str(e)}'}), 500

@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    # Envoie le fichier demandé
    return send_from_directory('uploads', filename)


if __name__ == '__main__':
    app.run(debug=True)
