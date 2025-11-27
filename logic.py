import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
import logging
logger = logging.getLogger(__name__)

def hex_to_rvb(hex_color: str) -> tuple:
    if hex_color is None:
        return None

    if hex_color.startswith("FF"):
        hex_color = hex_color[2:]

    try:
        r = int(hex_color[0:2], 16)
        v = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, v, b)
    except ValueError:
        logger.warning("Code couleur hex invalide : %s", hex_color, exc_info=True)
        return None

def extract_theme_colors(file_path: str) -> dict:
    theme_colors = {}

    if not os.path.exists(file_path):
        logger.error("Fichier thème introuvable : %s", file_path)
        return theme_colors

    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            if 'xl/theme/theme1.xml' not in zip_ref.namelist():
                logger.warning("Pas de thème dans le fichier : %s", file_path)
                return theme_colors
            with zip_ref.open('xl/theme/theme1.xml') as theme_file:
                tree = ET.parse(theme_file)
                root = tree.getroot()
                ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                for color_scheme in root.findall('a:themeElements/a:clrScheme', ns):
                    for i, color in enumerate(color_scheme):
                        rgb = color.find('a:srgbClr', ns)
                        if rgb is not None:
                            theme_colors[i] = rgb.attrib['val']
    except Exception:
        logger.error("Erreur lors de l'extraction des couleurs du thème", exc_info=True)

    return theme_colors

def get_implantation_colors(file_path: str, sheet_name: str) -> dict:
    if not os.path.exists(file_path):
        logger.error("Fichier source introuvable pour get_implantation_colors : %s", file_path)
        return {}

    try:
        theme_colors = extract_theme_colors(file_path)
        workbook = load_workbook(filename=file_path, data_only=True)
    except Exception:
        logger.error("Erreur lors de l'ouverture du fichier source : %s", file_path, exc_info=True)
        return {}

    try:
        if sheet_name not in workbook.sheetnames:
            logger.error("Feuille source introuvable : %s dans %s", sheet_name, file_path)
            return {}

        sheet = workbook[sheet_name]

        data_colors = {}
        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
            implantation = row[0].value
            nom = row[1].value
            prenom = row[2].value

            if implantation is None or nom is None or prenom is None:
                continue

            key = (implantation, nom, prenom)
            cell_a = row[0]

            if cell_a.fill and cell_a.fill.fill_type != "none":
                bg_color = cell_a.fill.fgColor
                rvb_color = None

                if bg_color.type == "rgb":
                    rvb_color = hex_to_rvb(bg_color.rgb)
                elif bg_color.type == "theme":
                    hex_color = theme_colors.get(bg_color.theme)
                    rvb_color = hex_to_rvb(hex_color)

                if rvb_color:
                    data_colors[key] = rvb_color

        return data_colors
    except Exception:
        logger.error("Erreur lors de l'extraction des couleurs d'implantation", exc_info=True)
        return {}
    finally:
        workbook.close()

def apply_colors_to_file2(file1_path: str, file1_sheet: str, file2_path: str, file2_sheet: str) -> None:
    if not os.path.exists(file2_path):
        logger.error("Fichier cible introuvable : %s", file2_path)
        return

    data_colors = get_implantation_colors(file1_path, file1_sheet)

    try:
        workbook = load_workbook(file2_path)
    except Exception:
        logger.error("Erreur lors de l'ouverture du fichier cible : %s", file2_path, exc_info=True)
        return

    try:
        if file2_sheet not in workbook.sheetnames:
            logger.error("Feuille cible introuvable : %s dans %s", file2_sheet, file2_path)
            return

        sheet = workbook[file2_sheet]

        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
            implantation = row[0].value
            nom = row[1].value
            prenom = row[2].value

            if implantation is None or nom is None or prenom is None:
                continue

            key = (implantation, nom, prenom)
            rvb_color = data_colors.get(key)

            if rvb_color:
                fill = PatternFill(
                    start_color=f"{rvb_color[0]:02X}{rvb_color[1]:02X}{rvb_color[2]:02X}",
                    end_color=f"{rvb_color[0]:02X}{rvb_color[1]:02X}{rvb_color[2]:02X}",
                    fill_type="solid"
                )
                for cell in row:
                    cell.fill = fill

        workbook.save(file2_path)
    except Exception:
        logger.error("Erreur lors de l'application des couleurs au fichier cible", exc_info=True)
    finally:
        workbook.close()

