from pathlib import Path
import math
import re
import shutil
from datetime import datetime
import unicodedata

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A5, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

try:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import RGBColor
    from docx.shared import Pt
except ModuleNotFoundError:
    Document = None
    WD_CELL_VERTICAL_ALIGNMENT = None
    WD_ALIGN_PARAGRAPH = None
    OxmlElement = None
    qn = None
    RGBColor = None
    Pt = None

FICHIER_EXCEL = "Notes.xlsx"
FEUILLE = 0
DOSSIER_SORTIE = "bulletins"
NOM_PDF = "bulletins_tous_eleves.pdf"
ANNEE = "ANNEE SCOLAIRE 2025-2026"
TITRE = "BULLETIN DU DEUXIEME TRIMESTRE"
CEB = "CEB DE NAMISSIGUIMA"
ECOLE = "ECOLE DE NAMISSIGUIMA B"
CLASSE = "CLASSE DE CP2 A"
COLONNE_NOM = "Nom & Prénom"
COLONNE_TOTAL = "Total"
COLONNE_MOYENNE = "Moyenne"
COLONNE_RANG = "Rang"
COLONNE_OBSERVATION = "Observation"
COLONNE_PHOTO = "Photo"
COLONNE_SEXE = "Sexe"
COLONNE_AGE = "Age"
COLONNE_SCOLARITE = "Scolarité"
COLONNES_PHOTO_POSSIBLES = ["Photo", "Photo eleve", "Photo élève", "Image", "Chemin photo"]
MATIERES = [
    {"libelle": "Exercices d'observation", "colonne": "Exercices d'obsevation"},
    {"libelle": "Expression orale", "colonne": "Expression\n orale"},
    {"libelle": "Récit / Chant", "colonne": "Recit/\nchant"},
    {"libelle": "Lecture", "colonne": "Lecture"},
    {"libelle": "Écriture", "colonne": "Écriture "},
    {"libelle": "Copie", "colonne": "Copie"},
    {"libelle": "Dictée", "colonne": "Dictée "},
    {"libelle": "Dessin", "colonne": "Dessin"},
    {"libelle": "Calcul", "colonne": "Calcul "},
    {"libelle": "EMC", "colonne": "EMC"},
    {"libelle": "TIC", "colonne": "TIC"},
]
BLEU = colors.HexColor("#1E0EF1")
BLEU_FONCE = colors.HexColor("#0640BD")
BLEU_CLAIR = colors.HexColor("#EAF3FA")
VERT = colors.HexColor("#168A53")
ROUGE = colors.HexColor("#C0392B")
OR = colors.HexColor("#009E42")
GRIS_TEXTE = colors.HexColor("#526272")
GRIS_CLAIR = colors.HexColor("#F5F7FA")
BLANC = colors.white

class ValidationErreur(Exception):
    pass

def normaliser_nom_colonne(nom):
    texte = str(nom).replace("\n", " ").replace("’", "'")
    texte = " ".join(texte.split()).strip().lower()
    texte = unicodedata.normalize("NFKD", texte)
    return "".join(c for c in texte if not unicodedata.combining(c))

def slugifier(texte):
    texte = re.sub(r"[^a-z0-9]+", "_", normaliser_nom_colonne(texte))
    return texte.strip("_") or "document"

def trouver_colonne(colonnes, nom_attendu):
    attendu = normaliser_nom_colonne(nom_attendu)
    for colonne in colonnes:
        if normaliser_nom_colonne(colonne) == attendu:
            return colonne
    return None

def lister_feuilles(fichier_excel):
    chemin = Path(fichier_excel)
    if not chemin.exists():
        raise ValidationErreur(f"Fichier Excel introuvable : {chemin.resolve()}")
    return pd.ExcelFile(chemin).sheet_names

def charger_notes(fichier_excel=FICHIER_EXCEL, feuille=FEUILLE):
    chemin = Path(fichier_excel)
    if not chemin.exists():
        raise ValidationErreur(f"Fichier Excel introuvable : {chemin.resolve()}")
    try:
        return pd.read_excel(chemin, sheet_name=feuille)
    except Exception as exc:
        raise ValidationErreur(f"Impossible de lire le fichier Excel : {exc}") from exc

def valider_donnees(df):
    erreurs, avertissements, correspondances = [], [], {}
    colonnes = list(df.columns)
    obligatoires = [COLONNE_NOM, COLONNE_TOTAL, COLONNE_MOYENNE, COLONNE_RANG, COLONNE_OBSERVATION]
    obligatoires.extend(m["colonne"] for m in MATIERES)
    for colonne in obligatoires:
        trouvee = trouver_colonne(colonnes, colonne)
        if trouvee is None:
            erreurs.append(f"Colonne obligatoire manquante : {colonne!r}")
        else:
            correspondances[colonne] = trouvee
    for colonne_photo in COLONNES_PHOTO_POSSIBLES:
        trouvee = trouver_colonne(colonnes, colonne_photo)
        if trouvee is not None:
            correspondances[COLONNE_PHOTO] = trouvee
            break
    if erreurs:
        raise ValidationErreur("\n".join(erreurs))
    if df.empty:
        raise ValidationErreur("Le fichier Excel ne contient aucun eleve.")
    colonnes_matieres = [correspondances[m["colonne"]] for m in MATIERES]
    notes = df[colonnes_matieres].apply(pd.to_numeric, errors="coerce")
    total = pd.to_numeric(df[correspondances[COLONNE_TOTAL]], errors="coerce")
    moyenne = pd.to_numeric(df[correspondances[COLONNE_MOYENNE]], errors="coerce")
    rang = pd.to_numeric(df[correspondances[COLONNE_RANG]], errors="coerce")
    for index, eleve in df.iterrows():
        ligne = index + 2
        nom = str(eleve[correspondances[COLONNE_NOM]]).strip()
        if not nom or nom.lower() == "nan":
            erreurs.append(f"Ligne {ligne} : nom d'eleve manquant.")
        for matiere, colonne_source in zip(MATIERES, colonnes_matieres):
            valeur = eleve[colonne_source]
            valeur_num = pd.to_numeric(pd.Series([valeur]), errors="coerce").iloc[0]
            if pd.isna(valeur) or pd.isna(valeur_num):
                avertissements.append(f"Ligne {ligne} ({nom}) : note vide ou non numerique pour {matiere['libelle']}.")
    if total.isna().any():
        erreurs.append("Total absent ou non numerique aux lignes : " + ", ".join(str(i + 2) for i in df.index[total.isna()]) + ".")
    if moyenne.isna().any():
        erreurs.append("Moyenne absente ou non numerique aux lignes : " + ", ".join(str(i + 2) for i in df.index[moyenne.isna()]) + ".")
    if rang.isna().any():
        erreurs.append("Rang absent ou non numerique aux lignes : " + ", ".join(str(i + 2) for i in df.index[rang.isna()]) + ".")
    total_calcule = notes.fillna(0).sum(axis=1)
    total_different = (total_calcule.round(6) != total.round(6)) & ~total.isna()
    if total_different.any():
        avertissements.append("Total different de la somme des notes aux lignes : " + ", ".join(str(i + 2) for i in df.index[total_different]) + ".")
    moyenne_calculee = total_calcule / len(MATIERES)
    moyenne_differente = (moyenne_calculee.round(6) != moyenne.round(6)) & ~moyenne.isna()
    if moyenne_differente.any():
        avertissements.append(f"Moyenne differente du calcul Total / {len(MATIERES)} aux lignes : " + ", ".join(str(i + 2) for i in df.index[moyenne_differente]) + ".")
    if erreurs:
        raise ValidationErreur("\n".join(erreurs))
    return correspondances, avertissements

def format_note(valeur):
    if pd.isna(valeur):
        return ""
    try:
        nombre = float(valeur)
    except (TypeError, ValueError):
        return str(valeur).strip()
    if math.isfinite(nombre) and nombre.is_integer():
        return str(int(nombre))
    return f"{nombre:g}".replace(".", ",")

def format_nombre(valeur, decimales=0):
    if pd.isna(valeur):
        return ""
    nombre = float(valeur)
    if decimales == 0:
        return str(int(nombre))
    return f"{nombre:.{decimales}f}".replace(".", ",")

def texte_ajuste(c, texte, x, y, largeur_max, police="Helvetica", taille=10, taille_min=7):
    texte = str(texte)
    while taille > taille_min and c.stringWidth(texte, police, taille) > largeur_max:
        taille -= 0.5
    c.setFont(police, taille)
    c.drawString(x, y, texte)

def entete_dict(annee=ANNEE, titre=TITRE, ceb=CEB, ecole=ECOLE, classe=CLASSE):
    return {"annee": annee, "titre": titre, "ceb": ceb, "ecole": ecole, "classe": classe}

def chemin_photo(eleve, correspondances, dossier_base):
    colonne_photo = correspondances.get(COLONNE_PHOTO)
    if not colonne_photo:
        return None
    valeur = eleve.get(colonne_photo, "")
    if pd.isna(valeur) or not str(valeur).strip():
        return None
    chemin = Path(str(valeur).strip())
    if not chemin.is_absolute():
        chemin = Path(dossier_base) / chemin
    return chemin if chemin.exists() else None

def verifier_photos(df, correspondances, dossier_base):
    avertissements = []
    colonne_photo = correspondances.get(COLONNE_PHOTO)
    if not colonne_photo:
        return avertissements
    for index, eleve in df.iterrows():
        valeur = eleve.get(colonne_photo, "")
        if pd.isna(valeur) or not str(valeur).strip():
            continue
        chemin = Path(str(valeur).strip())
        if not chemin.is_absolute():
            chemin = Path(dossier_base) / chemin
        if not chemin.exists():
            nom = str(eleve[correspondances[COLONNE_NOM]]).strip()
            avertissements.append(f"Ligne {index + 2} ({nom}) : photo introuvable ({chemin}).")
    return avertissements

def dessiner_photo(c, photo_path, x, y, largeur, hauteur):
    c.setStrokeColor(BLEU)
    c.roundRect(x, y, largeur, hauteur, 8, stroke=1, fill=0)
    if photo_path:
        try:
            c.drawImage(ImageReader(str(photo_path)), x + 3, y + 3, width=largeur - 6, height=hauteur - 6, preserveAspectRatio=True, anchor="c", mask="auto")
            return
        except Exception:
            pass
    c.setFillColor(BLEU_CLAIR)
    c.roundRect(x + 3, y + 3, largeur - 6, hauteur - 6, 6, stroke=0, fill=1)
    c.setFillColor(BLEU)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(x + largeur / 2, y + hauteur / 2 + 4, "PHOTO")
    c.setFont("Helvetica", 6.5)
    c.drawCentredString(x + largeur / 2, y + hauteur / 2 - 7, "optionnelle")

def dessiner_entete(c, largeur, hauteur, entete):
    c.setFillColor(BLEU_FONCE)
    c.rect(0, hauteur - 88, largeur, 88, stroke=0, fill=1)
    c.setFillColor(BLEU)
    c.rect(0, hauteur - 88, largeur, 9, stroke=0, fill=1)
    c.setFillColor(OR)
    c.rect(0, hauteur - 9, largeur, 9, stroke=0, fill=1)
    c.setFillColor(BLANC)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(28, hauteur - 26, entete["ceb"])
    c.drawRightString(largeur - 28, hauteur - 26, entete["annee"])
    c.setFont("Helvetica", 9)
    c.drawString(28, hauteur - 42, entete["ecole"])
    c.drawRightString(largeur - 28, hauteur - 42, entete["classe"])
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(largeur / 2, hauteur - 67, entete["titre"])
    c.setStrokeColor(BLEU_FONCE)
    c.roundRect(15, 15, largeur - 30, hauteur - 30, 10, stroke=1, fill=0)

def dessiner_carte_eleve(c, eleve, correspondances, dossier_photos, largeur, hauteur):
    y, x, w, h = hauteur - 158, 25, largeur - 50, 56
    c.setFillColor(BLANC)
    c.setStrokeColor(BLEU_CLAIR)
    c.roundRect(x, y, w, h, 8, stroke=1, fill=1)
    nom = str(eleve[correspondances[COLONNE_NOM]]).strip()
    c.setFillColor(GRIS_TEXTE)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(x + 12, y + h - 17, "ELEVE")
    c.setFillColor(BLEU_FONCE)
    texte_ajuste(c, nom, x + 12, y + h - 36, 245, "Helvetica-Bold", 13, 8)
    photo_path = chemin_photo(eleve, correspondances, dossier_photos)
    if photo_path:
        dessiner_photo(c, photo_path, largeur - 88, y - 6, 58, 68)
    return y - 24

def dessiner_tableau_notes(c, eleve, correspondances, largeur, y):
    x = 25
    c.setFillColor(BLEU)
    c.roundRect(x, y - 6, largeur - 50, 22, 5, stroke=0, fill=1)
    c.setFillColor(BLANC)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x + 10, y, "MATIERE")
    c.drawRightString(largeur - 35, y, "NOTE")
    y -= 22
    for index, matiere in enumerate(MATIERES):
        if index % 2 == 0:
            c.setFillColor(GRIS_CLAIR)
            c.roundRect(x, y - 5, largeur - 50, 17, 3, stroke=0, fill=1)
        c.setFillColor(GRIS_TEXTE)
        c.setFont("Helvetica", 9.5)
        c.drawString(x + 10, y, matiere["libelle"])
        c.setFillColor(BLEU_FONCE)
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(largeur - 37, y, format_note(eleve[correspondances[matiere["colonne"]]]))
        y -= 18
    return y

def dessiner_badge(c, x, y, largeur, titre, valeur, couleur):
    c.setFillColor(couleur)
    c.roundRect(x, y, largeur, 31, 7, stroke=0, fill=1)
    c.setFillColor(BLANC)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString(x + largeur / 2, y + 19, titre)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(x + largeur / 2, y + 6, str(valeur))

def dessiner_resultats(c, eleve, correspondances, largeur, y):
    y -= 30
    marge, espace = 25, 8
    w = (largeur - 2 * marge - 2 * espace) / 3
    dessiner_badge(c, marge, y, w, "TOTAL", format_nombre(eleve[correspondances[COLONNE_TOTAL]]), BLEU_FONCE)
    dessiner_badge(c, marge + w + espace, y, w, "MOYENNE", format_nombre(eleve[correspondances[COLONNE_MOYENNE]], 2), BLEU)
    dessiner_badge(c, marge + 2 * (w + espace), y, w, "RANG", format_nombre(eleve[correspondances[COLONNE_RANG]]), VERT)
    y -= 28
    observation = str(eleve[correspondances[COLONNE_OBSERVATION]]).strip()
    couleur = VERT if observation.lower() == "validé" else ROUGE
    c.setFillColor(BLANC)
    c.setStrokeColor(couleur)
    c.roundRect(25, y - 12, largeur - 50, 24, 6, stroke=1, fill=1)
    c.setFillColor(BLEU_FONCE)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(36, y - 3, "Observation")
    c.setFillColor(couleur)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(largeur / 2 + 45, y - 3, observation)

def dessiner_signatures(c, largeur):
    y = 44
    c.setFont("Helvetica", 9)
    c.setStrokeColor(BLEU_FONCE)
    c.setFillColor(GRIS_TEXTE)
    c.drawString(30, y + 15, "LES PROFESSEURS")
    c.line(30, y, 130, y)
    c.drawRightString(largeur - 30, y + 15, "LES PARENTS")
    c.line(largeur - 130, y, largeur - 30, y)

def sauvegarder_canvas(c, pdf_path):
    try:
        c.save()
    except PermissionError as exc:
        raise ValidationErreur(f"Impossible d'ecrire le PDF : {pdf_path}. Fermez le fichier s'il est deja ouvert, puis relancez.") from exc

def generer_pdf(df, correspondances, dossier_sortie=DOSSIER_SORTIE, nom_pdf=NOM_PDF, entete=None, dossier_photos="."):
    entete = entete or entete_dict()
    dossier_sortie = Path(dossier_sortie)
    dossier_sortie.mkdir(parents=True, exist_ok=True)
    pdf_path = dossier_sortie / nom_pdf
    c = canvas.Canvas(str(pdf_path), pagesize=A5)
    largeur, hauteur = A5
    for _, eleve in df.iterrows():
        dessiner_entete(c, largeur, hauteur, entete)
        y = dessiner_carte_eleve(c, eleve, correspondances, dossier_photos, largeur, hauteur)
        y = dessiner_tableau_notes(c, eleve, correspondances, largeur, y)
        dessiner_resultats(c, eleve, correspondances, largeur, y)
        dessiner_signatures(c, largeur)
        c.showPage()
    sauvegarder_canvas(c, pdf_path)
    return pdf_path

def generer_bulletins(fichier_excel=FICHIER_EXCEL, feuille=FEUILLE, dossier_sortie=DOSSIER_SORTIE, nom_pdf=NOM_PDF, annee=ANNEE, titre=TITRE, ceb=CEB, ecole=ECOLE, classe=CLASSE):
    df = charger_notes(fichier_excel, feuille)
    correspondances, avertissements = valider_donnees(df)
    dossier_photos = Path(fichier_excel).resolve().parent
    avertissements.extend(verifier_photos(df, correspondances, dossier_photos))
    pdf_path = generer_pdf(df, correspondances, dossier_sortie, nom_pdf, entete_dict(annee, titre, ceb, ecole, classe), dossier_photos)
    return {"pdf_path": pdf_path, "eleves": len(df), "pages": len(df), "avertissements": avertissements}

def generer_bulletins_trimestre(fichier_excel, feuille, libelle_trimestre, dossier_sortie, annee=ANNEE, ceb=CEB, ecole=ECOLE, classe=CLASSE):
    return generer_bulletins(fichier_excel, feuille, dossier_sortie, f"bulletins_{slugifier(libelle_trimestre)}.pdf", annee, f"BULLETIN DU {libelle_trimestre.upper()}", ceb, ecole, classe)

def extraire_resultats_trimestre(fichier_excel, feuille, code):
    df = charger_notes(fichier_excel, feuille)
    correspondances, avertissements = valider_donnees(df)
    lignes = []
    for _, eleve in df.iterrows():
        nom = str(eleve[correspondances[COLONNE_NOM]]).strip()
        lignes.append({"cle": normaliser_nom_colonne(nom), "nom": nom, f"total_{code}": pd.to_numeric(eleve[correspondances[COLONNE_TOTAL]], errors="coerce"), f"moyenne_{code}": pd.to_numeric(eleve[correspondances[COLONNE_MOYENNE]], errors="coerce"), f"rang_{code}": pd.to_numeric(eleve[correspondances[COLONNE_RANG]], errors="coerce"), f"observation_{code}": str(eleve[correspondances[COLONNE_OBSERVATION]]).strip()})
    return pd.DataFrame(lignes), avertissements

def construire_bilan_annuel(fichier_excel, selections):
    bilan, avertissements = None, []
    for code, feuille in selections:
        donnees, warns = extraire_resultats_trimestre(fichier_excel, feuille, code)
        avertissements.extend([f"{code} : {w}" for w in warns])
        bilan = donnees if bilan is None else bilan.merge(donnees, on=["cle", "nom"], how="outer")
    if bilan is None or bilan.empty:
        raise ValidationErreur("Aucune donnee trimestrielle valide pour le bilan annuel.")
    colonnes_moyennes = [f"moyenne_{code}" for code, _ in selections]
    bilan["moyenne_annuelle"] = bilan[colonnes_moyennes].mean(axis=1, skipna=True)
    bilan["rang_annuel"] = bilan["moyenne_annuelle"].rank(method="min", ascending=False).astype("Int64")
    bilan["observation_annuelle"] = bilan["moyenne_annuelle"].apply(lambda m: "Validé" if pd.notna(m) and m >= 5 else "Non validé")
    return bilan.sort_values(["rang_annuel", "nom"], na_position="last").reset_index(drop=True), avertissements

def libelle_periode_annuelle(code):
    correspondances = {
        "T1": "1er Trimestre",
        "1er Trimestre": "1er Trimestre",
        "PREMIER TRIMESTRE": "1er Trimestre",
        "T2": "2eme Trimestre",
        "2eme Trimestre": "2eme Trimestre",
        "DEUXIEME TRIMESTRE": "2eme Trimestre",
        "T3": "3eme Trimestre",
        "3eme Trimestre": "3eme Trimestre",
        "TROISIEME TRIMESTRE": "3eme Trimestre",
    }
    return correspondances.get(str(code).strip(), str(code).strip())

def decision_conseil_annuelle(moyenne, classe):
    classe_texte = str(classe or "").replace("CLASSE DE ", "").strip().upper()
    niveau = re.search(r"\b(CP1|CP2|CE1|CE2|CM1|CM2)\b", classe_texte)
    classe_actuelle = niveau.group(1) if niveau else (classe_texte or "CP2")
    classe_suivante = {
        "CP1": "CP2",
        "CP2": "CE1",
        "CE1": "CE2",
        "CE2": "CM1",
        "CM1": "CM2",
        "CM2": "6e",
    }.get(classe_actuelle, "classe supérieure")
    if pd.notna(moyenne) and moyenne >= 5:
        return f"Passe au {classe_suivante}"
    return f"Redouble au {classe_actuelle}"

def generer_bulletins_annuels(fichier_excel, selections, dossier_sortie, annee=ANNEE, ceb=CEB, ecole=ECOLE, classe=CLASSE):
    bilan, avertissements = construire_bilan_annuel(fichier_excel, selections)
    dossier_sortie = Path(dossier_sortie)
    dossier_sortie.mkdir(parents=True, exist_ok=True)
    pdf_path = dossier_sortie / "bulletins_annuels.pdf"
    c = canvas.Canvas(str(pdf_path), pagesize=A5)
    largeur, hauteur = A5
    codes = [code for code, _ in selections]
    for _, eleve in bilan.iterrows():
        dessiner_entete(c, largeur, hauteur, entete_dict(annee, "BULLETIN ANNUEL", ceb, ecole, classe))
        y = hauteur - 150
        c.setFillColor(BLANC); c.setStrokeColor(BLEU_CLAIR); c.roundRect(25, y, largeur - 50, 50, 8, stroke=1, fill=1)
        c.setFillColor(BLEU_FONCE); texte_ajuste(c, eleve["nom"], 37, y + 20, 300, "Helvetica-Bold", 13, 8)
        y -= 30
        c.setFillColor(BLEU); c.roundRect(25, y - 6, largeur - 50, 22, 5, stroke=0, fill=1)
        c.setFillColor(BLANC); c.setFont("Helvetica-Bold", 8.5)
        c.drawString(35, y, "PERIODE"); c.drawRightString(175, y, "MOYENNE"); c.drawRightString(250, y, "RANG"); c.drawRightString(largeur - 35, y, "OBSERVATION")
        y -= 24
        for index, code in enumerate(codes):
            if index % 2 == 0:
                c.setFillColor(GRIS_CLAIR); c.roundRect(25, y - 5, largeur - 50, 18, 3, stroke=0, fill=1)
            c.setFillColor(GRIS_TEXTE); c.setFont("Helvetica", 9.5)
            c.drawString(35, y, libelle_periode_annuelle(code)); c.drawRightString(175, y, format_nombre(eleve.get(f"moyenne_{code}"), 2)); c.drawRightString(250, y, format_nombre(eleve.get(f"rang_{code}"))); c.drawRightString(largeur - 35, y, str(eleve.get(f"observation_{code}", "")))
            y -= 20
        y -= 20; w = (largeur - 58) / 2
        dessiner_badge(c, 25, y, w, "MOYENNE ANNUELLE", format_nombre(eleve["moyenne_annuelle"], 2), BLEU)
        dessiner_badge(c, 33 + w, y, w, "RANG ANNUEL", format_nombre(eleve["rang_annuel"]), VERT)
        y -= 34; decision = decision_conseil_annuelle(eleve["moyenne_annuelle"], classe); couleur = VERT if str(decision).lower().startswith("passe") else ROUGE
        c.setStrokeColor(couleur); c.setFillColor(BLANC); c.roundRect(25, y - 18, largeur - 50, 30, 6, stroke=1, fill=1)
        c.setFillColor(BLEU_FONCE); c.setFont("Helvetica-Bold", 8.4); c.drawString(36, y, "Décision du Conseil des Professeurs :")
        c.setFillColor(couleur); c.setFont("Helvetica-Bold", 9.4); c.drawRightString(largeur - 36, y, decision)
        dessiner_signatures(c, largeur); c.showPage()
    sauvegarder_canvas(c, pdf_path)
    return {"pdf_path": pdf_path, "eleves": len(bilan), "pages": len(bilan), "avertissements": avertissements}

def generer_tableau_synoptique(fichier_excel, feuille, libelle_trimestre, dossier_sortie, annee=ANNEE, ceb=CEB, ecole=ECOLE, classe=CLASSE):
    df = charger_notes(fichier_excel, feuille)
    correspondances, avertissements = valider_donnees(df)
    dossier_sortie = Path(dossier_sortie); dossier_sortie.mkdir(parents=True, exist_ok=True)
    pdf_path = dossier_sortie / f"tableau_synoptique_{slugifier(libelle_trimestre)}.pdf"
    page_size = landscape(A4); c = canvas.Canvas(str(pdf_path), pagesize=page_size); largeur, hauteur = page_size
    lignes_par_page = 24; total_pages = max(1, math.ceil(len(df) / lignes_par_page))
    moyenne_classe = pd.to_numeric(df[correspondances[COLONNE_MOYENNE]], errors="coerce").mean()
    valides = df[correspondances[COLONNE_OBSERVATION]].astype(str).str.lower().str.contains("validé").sum()
    for page in range(total_pages):
        partie = df.iloc[page * lignes_par_page:(page + 1) * lignes_par_page]
        c.setFillColor(BLEU_FONCE); c.rect(0, hauteur - 72, largeur, 72, stroke=0, fill=1)
        c.setFillColor(BLANC); c.setFont("Helvetica-Bold", 14); c.drawString(32, hauteur - 30, f"TABLEAU SYNOPTIQUE - {libelle_trimestre.upper()}")
        c.setFont("Helvetica", 9); c.drawString(32, hauteur - 48, f"{ecole} | {classe} | {annee}"); c.drawRightString(largeur - 32, hauteur - 30, ceb); c.drawRightString(largeur - 32, hauteur - 48, f"Moyenne classe : {format_nombre(moyenne_classe, 2)} | Validés : {valides}")
        y = hauteur - 98; c.setFillColor(BLEU); c.roundRect(28, y - 8, largeur - 56, 24, 4, stroke=0, fill=1)
        c.setFillColor(BLANC); c.setFont("Helvetica-Bold", 9)
        for x, label in [(32, "N°"), (72, "Nom et prenom"), (370, "Total"), (450, "Moyenne"), (535, "Rang"), (610, "Observation")]: c.drawString(x, y, label)
        y -= 24
        for local_index, (_, eleve) in enumerate(partie.iterrows()):
            if local_index % 2 == 0: c.setFillColor(GRIS_CLAIR); c.rect(28, y - 6, largeur - 56, 19, stroke=0, fill=1)
            c.setFillColor(GRIS_TEXTE); c.setFont("Helvetica", 8.8)
            c.drawString(32, y, str(eleve.get("N°", page * lignes_par_page + local_index + 1)))
            texte_ajuste(c, str(eleve[correspondances[COLONNE_NOM]]).strip(), 72, y, 275, "Helvetica", 8.8, 6.5)
            c.drawRightString(410, y, format_nombre(eleve[correspondances[COLONNE_TOTAL]])); c.drawRightString(500, y, format_nombre(eleve[correspondances[COLONNE_MOYENNE]], 2)); c.drawRightString(565, y, format_nombre(eleve[correspondances[COLONNE_RANG]])); c.drawString(610, y, str(eleve[correspondances[COLONNE_OBSERVATION]]))
            y -= 20
        c.setFillColor(GRIS_TEXTE); c.setFont("Helvetica", 8); c.drawRightString(largeur - 32, 24, f"Page {page + 1}/{total_pages}"); c.showPage()
    sauvegarder_canvas(c, pdf_path)
    return {"pdf_path": pdf_path, "eleves": len(df), "pages": total_pages, "avertissements": avertissements}

def _docx_color(rgb):
    if not rgb or RGBColor is None:
        return None
    texte = str(rgb)
    if len(texte) == 8:
        texte = texte[-6:]
    if len(texte) != 6 or texte.upper() in {"000000", "FFFFFF"}:
        return None
    try:
        return RGBColor.from_string(texte)
    except ValueError:
        return None

def _docx_apply_font_name(run, font_name):
    if not font_name:
        return
    run.font.name = font_name
    if OxmlElement is None or qn is None:
        return
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        r_fonts.set(qn(attr), font_name)

def _docx_style_get(style, key, default=None):
    if hasattr(style, "get"):
        return style.get(key, default)
    return default

def _docx_set_cell(cell, text, size=7.2, bold=False, align=None, color=None, font_name=None, italic=False):
    if align is None:
        align = WD_ALIGN_PARAGRAPH.CENTER
    if isinstance(text, tuple):
        text, style = text
        if hasattr(style, "get"):
            color = _docx_style_get(style, "color") or color
            font_name = _docx_style_get(style, "font_name") or font_name
            bold = _docx_style_get(style, "bold", bold)
            italic = _docx_style_get(style, "italic", italic)
        else:
            color = style or color
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run(str(text))
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    _docx_apply_font_name(run, font_name)
    couleur = _docx_color(color)
    if couleur:
        run.font.color.rgb = couleur
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def _split_nom_prenoms(nom_complet):
    parties = str(nom_complet).strip().split()
    if not parties:
        return "", ""
    return parties[0], " ".join(parties[1:])

def _decision_fin_annee(moyenne):
    if pd.isna(moyenne):
        return ""
    return "Promu" if moyenne >= 5 else "Autorisé à redoubler"

def _couleur_cellule(cell):
    if cell.font and cell.font.color and cell.font.color.type == "rgb":
        rgb = cell.font.color.rgb
        if rgb and str(rgb) != "00000000":
            return str(rgb)
    return None

def _style_cellule(cell):
    font = cell.font
    if not font:
        return {}
    return {
        "color": _couleur_cellule(cell),
        "font_name": font.name,
        "bold": bool(font.bold),
        "italic": bool(font.italic),
    }

def _couleurs_feuille(fichier_excel, feuille):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(fichier_excel, data_only=True)
        ws = wb[feuille]
    except Exception:
        return {}

    headers = {normaliser_nom_colonne(cell.value): idx for idx, cell in enumerate(ws[1], start=1)}
    nom_idx = headers.get(normaliser_nom_colonne(COLONNE_NOM))
    if not nom_idx:
        return {}

    wanted = {
        "nom": COLONNE_NOM,
        "sexe": COLONNE_SEXE,
        "age": COLONNE_AGE,
        "scolarite": COLONNE_SCOLARITE,
        "moyenne": COLONNE_MOYENNE,
        "rang": COLONNE_RANG,
        "observation": COLONNE_OBSERVATION,
    }
    idxs = {key: headers.get(normaliser_nom_colonne(col)) for key, col in wanted.items()}
    result = {}
    for row in range(2, ws.max_row + 1):
        nom = ws.cell(row, nom_idx).value
        if not nom:
            continue
        cle = normaliser_nom_colonne(nom)
        result[cle] = {
            key: _style_cellule(ws.cell(row, idx))
            for key, idx in idxs.items()
            if idx
        }
    return result

def _donnees_proposition(fichier_excel, feuilles):
    if isinstance(feuilles, str):
        feuilles = [("T3", feuilles)]

    frames = {}
    avertissements = []
    for code, feuille in feuilles:
        df = charger_notes(fichier_excel, feuille)
        correspondances, warns = valider_donnees(df)
        couleurs = _couleurs_feuille(fichier_excel, feuille)
        avertissements.extend([f"{code} : {w}" for w in warns])
        sexe_col = trouver_colonne(df.columns, COLONNE_SEXE)
        age_col = trouver_colonne(df.columns, COLONNE_AGE)
        scolarite_col = trouver_colonne(df.columns, COLONNE_SCOLARITE)
        lignes = []
        for _, eleve in df.iterrows():
            nom = str(eleve[correspondances[COLONNE_NOM]]).strip()
            cle = normaliser_nom_colonne(nom)
            style = couleurs.get(cle, {})
            lignes.append({
                "cle": cle,
                "nom_complet": nom,
                f"moyenne_{code}": pd.to_numeric(eleve[correspondances[COLONNE_MOYENNE]], errors="coerce"),
                f"style_moyenne_{code}": style.get("moyenne", {}),
                f"observation_{code}": str(eleve[correspondances[COLONNE_OBSERVATION]]).strip(),
                "sexe": "" if sexe_col is None or pd.isna(eleve.get(sexe_col, "")) else str(eleve.get(sexe_col, "")).strip(),
                "age": "" if age_col is None or pd.isna(eleve.get(age_col, "")) else fmt_nombre_docx(eleve.get(age_col), 0),
                "scolarite": "" if scolarite_col is None or pd.isna(eleve.get(scolarite_col, "")) else fmt_nombre_docx(eleve.get(scolarite_col), 0),
                "style_ligne": style.get("nom") or style.get("sexe") or style.get("age") or style.get("scolarite") or style.get("moyenne") or style.get("observation") or {},
            })
        frames[code] = pd.DataFrame(lignes)

    if not frames:
        raise ValidationErreur("Aucune donnée disponible pour la proposition de fin d'année.")

    base_code = "T3" if "T3" in frames else list(frames.keys())[-1]
    bilan = frames[base_code].copy()
    for code, donnees in frames.items():
        if code == base_code:
            continue
        cols = ["cle", f"moyenne_{code}", f"style_moyenne_{code}", f"observation_{code}", "sexe", "age", "scolarite", "style_ligne"]
        bilan = bilan.merge(donnees[cols], on="cle", how="left", suffixes=("", f"_{code}"))
        for champ in ["sexe", "age", "scolarite"]:
            champ_cols = [col for col in bilan.columns if col == champ or col.startswith(f"{champ}_")]
            bilan[champ] = bilan[champ_cols].bfill(axis=1).iloc[:, 0].fillna("")
            bilan = bilan.drop(columns=[col for col in champ_cols if col != champ], errors="ignore")
        style_cols = [col for col in bilan.columns if col == "style_ligne" or col.startswith("style_ligne_")]
        bilan["style_ligne"] = bilan[style_cols].bfill(axis=1).iloc[:, 0]
        bilan = bilan.drop(columns=[col for col in style_cols if col != "style_ligne"], errors="ignore")

    if bilan.empty:
        raise ValidationErreur("Aucune donnée disponible pour la proposition de fin d'année.")

    codes = [code for code, _ in feuilles]
    for code in ["T1", "T2", "T3"]:
        col = f"moyenne_{code}"
        if col not in bilan.columns:
            bilan[col] = pd.NA
    bilan["moyenne_generale"] = bilan[[f"moyenne_{code}" for code in ["T1", "T2", "T3"]]].mean(axis=1, skipna=True)
    bilan["rang_general"] = bilan["moyenne_generale"].rank(method="min", ascending=False).astype("Int64")
    bilan["observation_finale"] = bilan.apply(
        lambda row: row.get("observation_T3", "") if pd.notna(row.get("observation_T3", "")) and str(row.get("observation_T3", "")).strip()
        else ("Validé" if pd.notna(row["moyenne_generale"]) and row["moyenne_generale"] >= 5 else "Non validé"),
        axis=1,
    )
    bilan = bilan.sort_values(["rang_general", "nom_complet"], na_position="last").reset_index(drop=True)
    return bilan, avertissements

def generer_proposition_fin_annee(
    fichier_excel,
    feuille,
    dossier_sortie,
    modele_docx=None,
    annee="2025-2026",
    ceb="NAMISSIGUIMA",
    ecole="NAMISSIGUIMA B",
    classe="CP2 A",
    lieu="NAMISSIGUIMA",
    date_document="19 mai 2026",
):
    if Document is None:
        raise ValidationErreur(
            "Le module python-docx n'est pas installé dans ce Python. "
            "Lancez l'application avec lancer_app.bat ou installez python-docx."
        )

    modele = Path(modele_docx) if modele_docx else Path(r"C:\Users\Abdoulaye MAIGA\Documents\PROPOSITION DE FIN D'ANNEE CP2 A.docx")
    if not modele.exists():
        raise ValidationErreur(f"Modèle Word introuvable : {modele}")

    dossier_sortie = Path(dossier_sortie)
    dossier_sortie.mkdir(parents=True, exist_ok=True)
    suffixe = datetime.now().strftime("%Y%m%d_%H%M%S")
    sortie = dossier_sortie / f"proposition_fin_annee_2026_{suffixe}.docx"
    shutil.copyfile(modele, sortie)

    bilan, avertissements = _donnees_proposition(fichier_excel, feuille)

    doc = Document(sortie)
    entetes = {
        0: f"Circonscription d’Education de Base de {ceb}\tAnnée scolaire {annee}",
        1: f"Ecole Primaire Publique de {ecole}                                                                       Classe : {classe}",
        6: f"Fait à {lieu} le {date_document}",
    }
    for index, texte in entetes.items():
        if index < len(doc.paragraphs):
            paragraph = doc.paragraphs[index]
            paragraph.text = texte
            for run in paragraph.runs:
                run.font.size = Pt(10)
                run.bold = index in (0, 1)

    lignes = []
    for index, eleve in bilan.iterrows():
        nom, prenoms = _split_nom_prenoms(eleve["nom_complet"])
        moyenne = eleve["moyenne_generale"]
        style_ligne = eleve.get("style_ligne", {})
        if not isinstance(style_ligne, dict):
            style_ligne = {}
        lignes.append([
            (index + 1, style_ligne),
            (nom, style_ligne),
            (prenoms, style_ligne),
            (eleve.get("sexe", ""), style_ligne),
            (eleve.get("age", ""), style_ligne),
            (eleve.get("scolarite", ""), style_ligne),
            (fmt_nombre_docx(eleve.get("moyenne_T1")), eleve.get("style_moyenne_T1") or style_ligne),
            (fmt_nombre_docx(eleve.get("moyenne_T2")), eleve.get("style_moyenne_T2") or style_ligne),
            (fmt_nombre_docx(eleve.get("moyenne_T3")), eleve.get("style_moyenne_T3") or style_ligne),
            (fmt_nombre_docx(moyenne), style_ligne),
            (fmt_nombre_docx(eleve["rang_general"], 0), style_ligne),
            (_decision_fin_annee(moyenne), style_ligne),
            (str(eleve.get("observation_finale", "")).strip(), style_ligne),
        ])

    ligne_index = 0
    for table_index in [0, 1, 2, 3]:
        if table_index >= len(doc.tables):
            break
        table = doc.tables[table_index]
        start_row = 2 if table_index == 0 else 0
        for row_index in range(start_row, len(table.rows)):
            if ligne_index >= len(lignes):
                break
            for col_index, valeur in enumerate(lignes[ligne_index]):
                _docx_set_cell(table.cell(row_index, col_index), valeur, size=6.6 if col_index in [1, 2, 11, 12] else 7)
            ligne_index += 1
        if ligne_index >= len(lignes):
            break

    if ligne_index < len(lignes) and len(doc.tables) > 3:
        table = doc.tables[3]
        while ligne_index < len(lignes):
            row = table.add_row()
            for col_index, valeur in enumerate(lignes[ligne_index]):
                _docx_set_cell(row.cells[col_index], valeur, size=6.6 if col_index in [1, 2, 11, 12] else 7)
            ligne_index += 1

    effectif = len(bilan)
    promus = int((bilan["moyenne_generale"] >= 5).sum())
    redoublants = int((bilan["moyenne_generale"] < 5).sum())

    if len(doc.tables) > 4:
        recap = doc.tables[4]
        _docx_set_cell(recap.cell(4, 0), classe.split()[0] if classe else "Classe", size=7.5, bold=True)
        for col, valeur in [(3, effectif), (6, effectif), (9, promus), (16, redoublants), (22, 0), (25, 0), (28, 0)]:
            if col < len(recap.columns):
                _docx_set_cell(recap.cell(4, col), valeur, size=7.5, bold=True)

    if len(doc.tables) > 5:
        inspecteur = doc.tables[5]
        _docx_set_cell(inspecteur.cell(3, 2), promus, size=8, bold=True)
        _docx_set_cell(inspecteur.cell(3, 5), redoublants, size=8, bold=True)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.space_before = Pt(0)
                    for run in paragraph.runs:
                        if run.font.size is None:
                            run.font.size = Pt(7)

    doc.save(sortie)
    return {
        "docx_path": sortie,
        "eleves": effectif,
        "pages": "",
        "promus": promus,
        "redoublants": redoublants,
        "avertissements": avertissements,
    }

def fmt_nombre_docx(valeur, decimales=2):
    if pd.isna(valeur):
        return ""
    if decimales == 0:
        return str(int(valeur))
    return f"{float(valeur):.{decimales}f}".replace(".", ",")

def main():
    try:
        resultat = generer_bulletins()
    except ValidationErreur as exc:
        print("Generation arretee :"); print(exc); raise SystemExit(1)
    if resultat["avertissements"]:
        print("Avertissements non bloquants :")
        for avertissement in resultat["avertissements"]: print(f"- {avertissement}")
    print("Generation terminee."); print(f"Eleves traites : {resultat['eleves']}"); print(f"Pages generees : {resultat['pages']}"); print(f"PDF : {resultat['pdf_path'].resolve()}")

if __name__ == "__main__":
    main()
