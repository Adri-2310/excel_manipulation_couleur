import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def hex_to_rvb(hex_color: str) -> tuple:
    """
    Convertit un code couleur hexadécimal en tuple RVB.

    Args:
        hex_color (str): Code couleur hexadécimal (ex: "FFFF0000").

    Returns:
        tuple: Valeurs RVB sous forme de tuple (R, V, B) ou None si la couleur est invalide.
    """
    if hex_color is None:
        return None

    if hex_color.startswith("FF"):
        hex_color = hex_color[2:]

    try:
        r = int(hex_color[0:2], 16)
        v = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, v, b)
    except:
        return None

def extract_theme_colors(file_path: str) -> dict:
    """
    Extrait les couleurs du thème d'un fichier Excel.

    Args:
        file_path (str): Chemin vers le fichier Excel.

    Returns:
        dict: Dictionnaire des couleurs du thème (clé : index, valeur : code hexadécimal).
    """
    theme_colors = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            if 'xl/theme/theme1.xml' in zip_ref.namelist():
                with zip_ref.open('xl/theme/theme1.xml') as theme_file:
                    tree = ET.parse(theme_file)
                    root = tree.getroot()
                    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

                    for color_scheme in root.findall('a:themeElements/a:clrScheme', ns):
                        for i, color in enumerate(color_scheme):
                            rgb = color.find('a:srgbClr', ns)
                            if rgb is not None:
                                theme_colors[i] = rgb.attrib['val']
    except Exception as e:
        print(f"Erreur lors de l'extraction des couleurs du thème : {e}")

    return theme_colors

def get_implantation_colors(file_path: str, sheet_name: str) -> dict:
    """
    Récupère les couleurs RVB associées aux combinaisons Implantation, Nom, Prénom dans le fichier 1.

    Args:
        file_path (str): Chemin vers le fichier Excel.
        sheet_name (str): Nom de la feuille.

    Returns:
        dict: Dictionnaire associant chaque combinaison (Implantation, Nom, Prénom) à sa couleur RVB.
    """
    theme_colors = extract_theme_colors(file_path)
    workbook = load_workbook(filename=file_path, data_only=True)
    sheet = workbook[sheet_name]

    # Utiliser un tuple (Implantation, Nom, Prénom) comme clé
    data_colors = {}

    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
        implantation = row[0].value
        nom = row[1].value
        prenom = row[2].value

        if implantation is None or nom is None or prenom is None:
            continue

        key = (implantation, nom, prenom)

        cell_a = row[0]
        if cell_a.fill.fill_type != "none":
            bg_color = cell_a.fill.fgColor
            if bg_color.type == "rgb":
                rvb_color = hex_to_rvb(bg_color.rgb)
                data_colors[key] = rvb_color
            elif bg_color.type == "theme":
                hex_color = theme_colors.get(bg_color.theme, None)
                rvb_color = hex_to_rvb(hex_color)
                data_colors[key] = rvb_color

    return data_colors

def apply_colors_to_file2(file1_path: str, file1_sheet: str, file2_path: str, file2_sheet: str) -> None:
    """
    Applique les couleurs du fichier 1 au fichier 2 pour les combinaisons Implantation, Nom, Prénom correspondantes,
    sur les colonnes A, B et C.

    Args:
        file1_path (str): Chemin vers le fichier 1 (source des couleurs).
        file1_sheet (str): Nom de la feuille dans le fichier 1.
        file2_path (str): Chemin vers le fichier 2 (destination des couleurs).
        file2_sheet (str): Nom de la feuille dans le fichier 2.
    """
    # Récupérer les couleurs du fichier 1
    data_colors = get_implantation_colors(file1_path, file1_sheet)

    # Charger le fichier 2 pour appliquer les couleurs
    workbook = load_workbook(file2_path)
    sheet = workbook[file2_sheet]

    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
        implantation = row[0].value
        nom = row[1].value
        prenom = row[2].value

        if implantation is None or nom is None or prenom is None:
            continue

        key = (implantation, nom, prenom)

        if key in data_colors:
            rvb_color = data_colors[key]
            if rvb_color:
                # Créer un remplissage avec la couleur RVB
                fill = PatternFill(
                    start_color=f"{rvb_color[0]:02X}{rvb_color[1]:02X}{rvb_color[2]:02X}",
                    end_color=f"{rvb_color[0]:02X}{rvb_color[1]:02X}{rvb_color[2]:02X}",
                    fill_type="solid"
                )
                # Appliquer la couleur aux colonnes A, B et C
                for cell in row:
                    cell.fill = fill

    # Sauvegarder les modifications
    workbook.save(file2_path)

