from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
import pandas as pd
import os
import traceback

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
            message = f"Traitement terminé. Fichier généré : {os.path.basename(file_target)}"
            return render_template(
                'result.html',
                message=message,
                chemin_fichier_modifie=file_target
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

@app.route('/download/<filename>')
def download(filename):
    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        sheet_name = request.args.get('sheet_name')
        if not sheet_name:
            return "Aucune feuille spécifiée pour le téléchargement.", 400

        df = pd.read_excel(filepath, sheet_name=sheet_name)

        if 'Total' not in df.columns:
            df['Total'] = df.sum(axis=1)

        modified_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f'modified_{filename}')
        with pd.ExcelWriter(modified_filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        return send_file(
            modified_filepath,
            as_attachment=True,
            download_name=f'modified_{filename}'
        )
    except Exception as e:
        print(f"Erreur dans download: {traceback.format_exc()}")
        return f"Une erreur est survenue : {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
