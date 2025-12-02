# tests/test_app.py
import io
import os
import pytest
from app import app

@pytest.fixture
def client(tmp_path, monkeypatch):
    """
    Client de test Flask avec un dossier uploads temporaire.
    """
    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()

    app.config['UPLOAD_FOLDER'] = str(upload_dir)
    app.config['TESTING'] = True

    with app.test_client() as client:
        yield client

def test_index_get(client):
    resp = client.get('/')
    assert resp.status_code == 200
    assert b"Manipulation de fichiers Excel" in resp.data  # texte de index.html [3]

def test_index_post_sans_fichier_source(client):
    # POST sans champ "file" → doit renvoyer 400 selon ton code [1]
    resp = client.post('/', data={})
    assert resp.status_code == 400
    assert b"Aucun fichier source t\xc3\xa9l\xc3\xa9vers\xc3\xa9" in resp.data

def test_index_post_fichier_source_nom_vide(client):
    data = {
        'file': (io.BytesIO(b""), ''),  # nom vide
    }
    resp = client.post('/', data=data, content_type='multipart/form-data')
    assert resp.status_code == 400
    assert b"Nom de fichier source vide" in resp.data

def test_get_sheets_sans_fichier(client):
    resp = client.post('/get_sheets')
    assert resp.status_code == 400
    assert resp.get_json()['error'] == 'Aucun fichier téléversé'

def test_get_sheets_nom_vide(client):
    data = {
        'file': (io.BytesIO(b""), ''),  # nom vide
    }
    resp = client.post('/get_sheets', data=data, content_type='multipart/form-data')
    assert resp.status_code == 400
    assert resp.get_json()['error'] == 'Nom de fichier vide'