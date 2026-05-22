from pathlib import Path
import sqlite3

import pandas as pd

import generer_bulletins as bulletins


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "ecole_local.db"
EXPORTS = ROOT / "exports"

TRIMESTRES = ["1er Trimestre", "2eme Trimestre", "3eme Trimestre"]


def connect():
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=OFF")
    conn.execute("PRAGMA synchronous=OFF")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                classe TEXT NOT NULL,
                nom_prenom TEXT NOT NULL,
                sexe TEXT,
                age TEXT,
                scolarite TEXT,
                photo TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(classe, nom_prenom)
            );

            CREATE TABLE IF NOT EXISTS grades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                trimestre TEXT NOT NULL,
                matiere TEXT NOT NULL,
                note TEXT,
                FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                UNIQUE(student_id, trimestre, matiere)
            );
            """
        )


def _row_to_dict(row):
    return dict(row) if row is not None else None


def classes():
    init_db()
    with connect() as conn:
        rows = conn.execute("SELECT DISTINCT classe FROM students ORDER BY classe").fetchall()
    return [row["classe"] for row in rows]


def students(classe=None):
    init_db()
    query = "SELECT * FROM students"
    params = []
    if classe:
        query += " WHERE classe = ?"
        params.append(classe)
    query += " ORDER BY nom_prenom"
    with connect() as conn:
        return [_row_to_dict(row) for row in conn.execute(query, params).fetchall()]


def save_student(data):
    init_db()
    classe = str(data.get("classe") or "").strip()
    nom = str(data.get("nom_prenom") or "").strip()
    if not classe:
        raise bulletins.ValidationErreur("La classe est obligatoire.")
    if not nom:
        raise bulletins.ValidationErreur("Le nom et prénom sont obligatoires.")

    values = {
        "classe": classe,
        "nom_prenom": nom,
        "sexe": str(data.get("sexe") or "").strip(),
        "age": str(data.get("age") or "").strip(),
        "scolarite": str(data.get("scolarite") or "").strip(),
        "photo": str(data.get("photo") or "").strip(),
    }
    student_id = data.get("id")
    with connect() as conn:
        if student_id:
            conn.execute(
                """
                UPDATE students
                SET classe = :classe, nom_prenom = :nom_prenom, sexe = :sexe,
                    age = :age, scolarite = :scolarite, photo = :photo,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = :id
                """,
                {**values, "id": student_id},
            )
        else:
            conn.execute(
                """
                INSERT INTO students(classe, nom_prenom, sexe, age, scolarite, photo)
                VALUES(:classe, :nom_prenom, :sexe, :age, :scolarite, :photo)
                ON CONFLICT(classe, nom_prenom) DO UPDATE SET
                    sexe = excluded.sexe,
                    age = excluded.age,
                    scolarite = excluded.scolarite,
                    photo = excluded.photo,
                    updated_at = CURRENT_TIMESTAMP
                """,
                values,
            )
        row = conn.execute(
            "SELECT * FROM students WHERE classe = ? AND nom_prenom = ?",
            (classe, nom),
        ).fetchone()
    return _row_to_dict(row)


def import_excel(fichier_excel, selections, classe):
    init_db()
    classe = str(classe or "").replace("CLASSE DE ", "").strip() or "Classe"
    imported_students = set()
    imported_notes = 0

    with connect() as conn:
        for trimestre, feuille in selections.items():
            if not feuille:
                continue
            df = bulletins.charger_notes(fichier_excel, feuille)
            correspondances, _ = bulletins.valider_donnees(df)
            sexe_col = bulletins.trouver_colonne(df.columns, bulletins.COLONNE_SEXE)
            age_col = bulletins.trouver_colonne(df.columns, bulletins.COLONNE_AGE)
            scolarite_col = bulletins.trouver_colonne(df.columns, bulletins.COLONNE_SCOLARITE)
            photo_col = correspondances.get(bulletins.COLONNE_PHOTO)

            for _, eleve in df.iterrows():
                nom = str(eleve[correspondances[bulletins.COLONNE_NOM]]).strip()
                if not nom or nom.lower() == "nan":
                    continue
                values = {
                    "classe": classe,
                    "nom_prenom": nom,
                    "sexe": "" if sexe_col is None or pd.isna(eleve.get(sexe_col, "")) else str(eleve.get(sexe_col, "")).strip(),
                    "age": "" if age_col is None or pd.isna(eleve.get(age_col, "")) else bulletins.fmt_nombre_docx(eleve.get(age_col), 0),
                    "scolarite": "" if scolarite_col is None or pd.isna(eleve.get(scolarite_col, "")) else bulletins.fmt_nombre_docx(eleve.get(scolarite_col), 0),
                    "photo": "" if photo_col is None or pd.isna(eleve.get(photo_col, "")) else str(eleve.get(photo_col, "")).strip(),
                }
                conn.execute(
                    """
                    INSERT INTO students(classe, nom_prenom, sexe, age, scolarite, photo)
                    VALUES(:classe, :nom_prenom, :sexe, :age, :scolarite, :photo)
                    ON CONFLICT(classe, nom_prenom) DO UPDATE SET
                        sexe = excluded.sexe,
                        age = excluded.age,
                        scolarite = excluded.scolarite,
                        photo = excluded.photo,
                        updated_at = CURRENT_TIMESTAMP
                    """,
                    values,
                )
                student = conn.execute(
                    "SELECT id FROM students WHERE classe = ? AND nom_prenom = ?",
                    (classe, nom),
                ).fetchone()
                imported_students.add(student["id"])
                for matiere in bulletins.MATIERES:
                    source = correspondances[matiere["colonne"]]
                    note = eleve.get(source, "")
                    note_text = "" if pd.isna(note) else str(note).strip()
                    conn.execute(
                        """
                        INSERT INTO grades(student_id, trimestre, matiere, note)
                        VALUES(?, ?, ?, ?)
                        ON CONFLICT(student_id, trimestre, matiere) DO UPDATE SET note = excluded.note
                        """,
                        (student["id"], trimestre, matiere["colonne"], note_text),
                    )
                    imported_notes += 1
    return {"students": len(imported_students), "notes": imported_notes, "classe": classe}


def notes_for_class(classe, trimestre):
    init_db()
    eleves = students(classe)
    with connect() as conn:
        for eleve in eleves:
            rows = conn.execute(
                "SELECT matiere, note FROM grades WHERE student_id = ? AND trimestre = ?",
                (eleve["id"], trimestre),
            ).fetchall()
            notes = {row["matiere"]: row["note"] for row in rows}
            eleve["notes"] = {matiere["colonne"]: notes.get(matiere["colonne"], "") for matiere in bulletins.MATIERES}
    return eleves


def save_notes(classe, trimestre, rows):
    init_db()
    saved = 0
    with connect() as conn:
        for row in rows:
            student_id = row.get("id")
            if not student_id:
                continue
            exists = conn.execute(
                "SELECT id FROM students WHERE id = ? AND classe = ?",
                (student_id, classe),
            ).fetchone()
            if not exists:
                continue
            notes = row.get("notes", {})
            for matiere in bulletins.MATIERES:
                note = str(notes.get(matiere["colonne"], "")).strip()
                conn.execute(
                    """
                    INSERT INTO grades(student_id, trimestre, matiere, note)
                    VALUES(?, ?, ?, ?)
                    ON CONFLICT(student_id, trimestre, matiere) DO UPDATE SET note = excluded.note
                    """,
                    (student_id, trimestre, matiere["colonne"], note),
                )
                saved += 1
    return {"saved": saved}


def export_excel(classe):
    init_db()
    EXPORTS.mkdir(exist_ok=True)
    output = EXPORTS / f"{bulletins.slugifier(classe)}_base_notes.xlsx"
    matiere_cols = [matiere["colonne"] for matiere in bulletins.MATIERES]
    all_students = students(classe)
    if not all_students:
        raise bulletins.ValidationErreur(f"Aucun élève inscrit pour la classe {classe}.")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for trimestre in TRIMESTRES:
            rows = []
            notes_rows = notes_for_class(classe, trimestre)
            notes_by_id = {item["id"]: item["notes"] for item in notes_rows}
            for index, eleve in enumerate(all_students, start=1):
                eleve_notes = notes_by_id.get(eleve["id"], {})
                row = {
                    "N°": index,
                    bulletins.COLONNE_NOM: eleve["nom_prenom"],
                    bulletins.COLONNE_SEXE: eleve.get("sexe", ""),
                    bulletins.COLONNE_AGE: eleve.get("age", ""),
                    bulletins.COLONNE_SCOLARITE: eleve.get("scolarite", ""),
                }
                numeric_notes = []
                for col in matiere_cols:
                    value = eleve_notes.get(col, "")
                    row[col] = value
                    numeric_notes.append(pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0])
                total = sum(0 if pd.isna(value) else float(value) for value in numeric_notes)
                moyenne = total / len(bulletins.MATIERES)
                row[bulletins.COLONNE_TOTAL] = total
                row[bulletins.COLONNE_MOYENNE] = moyenne
                row[bulletins.COLONNE_RANG] = ""
                row[bulletins.COLONNE_OBSERVATION] = "Validé" if moyenne >= 5 else "Non validé"
                rows.append(row)

            df = pd.DataFrame(rows)
            df[bulletins.COLONNE_RANG] = df[bulletins.COLONNE_MOYENNE].rank(method="min", ascending=False).astype(int)
            columns = ["N°", bulletins.COLONNE_NOM, bulletins.COLONNE_SEXE, bulletins.COLONNE_AGE, bulletins.COLONNE_SCOLARITE]
            columns.extend(matiere_cols)
            columns.extend([bulletins.COLONNE_TOTAL, bulletins.COLONNE_MOYENNE, bulletins.COLONNE_RANG, bulletins.COLONNE_OBSERVATION])
            df[columns].to_excel(writer, sheet_name=trimestre, index=False)
    return output
