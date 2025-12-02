from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import time
import glob
import threading
from flask import send_from_directory
import logging
from logging.handlers import RotatingFileHandler

# <-- importe ton script de logique ici (adapter le nom du fichier / fonction) -->
from logic import apply_colors_to_file2

app = Flask(__name__)

# Chemin de base du projet
basedir = os.path.dirname(os.path.abspath(__file__))

# Dossier pour les uploads
app.config['UPLOAD_FOLDER'] = os.path.join(basedir, 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# ------------------------------
#   CONFIGURATION DES LOGS
# ------------------------------
log_dir = os.path.join(basedir, 'log')
os.makedirs(log_dir, exist_ok=True)

log_path = os.path.join(log_dir, 'app.log')

handler = RotatingFileHandler(
    log_path,
    maxBytes=1_000_000,   # 1 Mo par fichier
    backupCount=5         # garde 5 anciens fichiers
)
formatter = logging.Formatter(
    '[%(asctime)s] %(levelname)s in %(module)s: %(message)s'
)
handler.setFormatter(formatter)

app.logger.setLevel(logging.INFO)
app.logger.addHandler(handler)


# ------------------------------
#   ROUTE PRINCIPALE
# ------------------------------
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template('index.html', table=None, filename=None, sheet_name=None)

    # POST = traitement
    try:
        # ----- FICHIER SOURCE (OBLIGATOIRE) -----
        if 'file' not in request.files:
            app.logger.warning("Tentative sans fichier source")
            return "Aucun fichier source téléversé.", 400

        file_source = request.files['file']
        if file_source.filename == '':
            app.logger.warning("Nom de fichier source vide")
            return "Nom de fichier source vide.", 400

        source_path = os.path.join(app.config['UPLOAD_FOLDER'], file_source.filename)
        file_source.save(source_path)

        sheet_source = request.form.get('sheet_name')
        if not sheet_source:
            app.logger.warning("Aucune feuille source sélectionnée")
            return "Aucune feuille source sélectionnée.", 400

        # ----- FICHIER CIBLE (OBLIGATOIRE) -----
        file_target = request.files.get('file_target')
        if not file_target or file_target.filename == '':
            app.logger.warning("Aucun fichier cible téléversé")
            return "Aucun fichier cible téléversé.", 400

        target_path = os.path.join(app.config['UPLOAD_FOLDER'], file_target.filename)
        file_target.save(target_path)

        sheet_target = request.form.get('sheet_name_target')
        if not sheet_target:
            app.logger.warning("Aucune feuille cible sélectionnée")
            return "Aucune feuille cible sélectionnée.", 400

        app.logger.info(
            "Nouveau traitement : source=%s, cible=%s, sheet_source=%s, sheet_target=%s",
            file_source.filename, file_target.filename, sheet_source, sheet_target
        )

        # ----- APPEL DE LA LOGIQUE METIER -----
        try:
            apply_colors_to_file2(
                file1_path=source_path,
                file1_sheet=sheet_source,
                file2_path=target_path,
                file2_sheet=sheet_target
            )
        except Exception:
            # On log l’erreur complète pour le debug
            app.logger.error("Erreur dans apply_colors_to_file2", exc_info=True)
            return "Erreur lors du traitement des fichiers Excel.", 500

        message = "Traitement terminé."
        return render_template(
            'resultat.html',
            message=message,
            chemin_fichier_modifie=os.path.basename(target_path),
        )

    except Exception:
        app.logger.error("Erreur inattendue dans la vue index", exc_info=True)
        return "Une erreur technique est survenue. Veuillez réessayer plus tard.", 500


# ------------------------------
#   VIDAGE DU DOSSIER UPLOADS
# ------------------------------
def vider_dossier_uploads(delai=60):
    """Vide tous les fichiers du dossier 'uploads' après un délai donné (en secondes)."""
    def vider():
        time.sleep(delai)
        try:
            fichiers = glob.glob(os.path.join(app.config['UPLOAD_FOLDER'], '*'))
            for fichier in fichiers:
                if os.path.isfile(fichier):
                    os.remove(fichier)
            app.logger.info("Dossier %s vidé avec succès.", app.config['UPLOAD_FOLDER'])
        except Exception:
            app.logger.error("Erreur lors du vidage du dossier uploads", exc_info=True)

    thread = threading.Thread(target=vider, daemon=True)
    thread.start()


@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    try:
        if 'file' not in request.files:
            app.logger.warning("get_sheets appelé sans fichier")
            return jsonify({'error': 'Aucun fichier téléversé'}), 400

        file = request.files['file']
        if file.filename == '':
            app.logger.warning("get_sheets : nom de fichier vide")
            return jsonify({'error': 'Nom de fichier vide'}), 400

        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        app.logger.info("Feuilles trouvées dans %s : %s", file.filename, sheet_names)
        return jsonify({'sheet_names': sheet_names})
    except Exception:
        app.logger.error("Erreur dans get_sheets", exc_info=True)
        return jsonify({'error': 'Erreur lors de la lecture du fichier.'}), 500

# ------------------------------
#   ROUTE DE TELECHARGEMENT
# ------------------------------
@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    if not os.path.exists(file_path):
        app.logger.warning("Fichier demandé introuvable : %s", file_path)
        return "Fichier non trouvé.", 404

    try:
        response = send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

        # Lancer un thread pour supprimer les fichiers après un délai
        vider_dossier_uploads()

        app.logger.info("Fichier envoyé : %s", file_path)
        return response
    except Exception:
        app.logger.error("Erreur lors de l'envoi du fichier %s", file_path, exc_info=True)
        return "Une erreur est survenue lors du téléchargement.", 500


# ------------------------------
#   HANDLERS D'ERREURS GLOBAUX
# ------------------------------
from werkzeug.exceptions import HTTPException

@app.errorhandler(404)
def not_found(e):
    app.logger.warning("404 sur l'URL : %s", request.path)
    return "Page non trouvée.", 404

@app.errorhandler(500)
def internal_error(e):
    app.logger.error("Erreur interne 500", exc_info=True)
    return "Erreur interne du serveur.", 500

@app.errorhandler(Exception)
def handle_exception(e):
    # Laisser passer les erreurs HTTP déjà gérées
    if isinstance(e, HTTPException):
        return e
    app.logger.error("Exception non gérée", exc_info=True)
    return "Une erreur inattendue est survenue.", 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)