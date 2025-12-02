# tests/test_logic.py
from logic import _find_columns_by_header, apply_colors_to_file2

def test_find_columns_by_header_ok(monkeypatch):
    """
    Vérifie qu'on trouve bien les index des colonnes par leur entête.
    """

    class FakeSheet:
        def iter_rows(self, min_row, max_row, values_only=False):
            # Une seule ligne d'en-tête
            # A: "Implantation", B: "Nom", C: "Prénom"
            yield ("Implantation", "Nom", "Prénom")

    sheet = FakeSheet()
    headers = {
        "implantation": ["implantation"],
        "nom": ["nom"],
        "prenom": ["prénom", "prenom"],
    }

    indices = _find_columns_by_header(sheet, headers)
    assert indices["implantation"] == 0
    assert indices["nom"] == 1
    assert indices["prenom"] == 2

def test_apply_colors_to_file2_with_missing_files(tmp_path):
    """
    Grâce aux vérifications os.path.exists, cette fonction ne doit pas lever d'erreur
    même si les fichiers n'existent pas encore.
    """
    file1 = tmp_path / "source.xlsx"
    file2 = tmp_path / "cible.xlsx"

    # On n'écrit volontairement aucun fichier, ils n'existent pas
    # La fonction doit simplement logger une erreur et retourner
    apply_colors_to_file2(str(file1), "Feuil1", str(file2), "Feuil1")
    # Si aucune exception n'est levée, le test est considéré comme OK