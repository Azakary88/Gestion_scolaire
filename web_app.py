from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse
import cgi
import json
import mimetypes
import traceback

import generer_bulletins as bulletins
import school_db


ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
OUTPUTS = ROOT / "bulletins"


HTML = r"""<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Gestion scolaire</title>
  <style>
    :root {
      --ink: #1f2933;
      --muted: #667085;
      --line: #d9e2ec;
      --soft: #f4f7fb;
      --blue: #1f4e79;
      --blue-dark: #173b5c;
      --gold: #d9a441;
      --green: #168a53;
      --red: #c0392b;
      --white: #fff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      color: var(--ink);
      background: #eef3f8;
    }
    header {
      background: var(--blue-dark);
      color: var(--white);
      padding: 18px 28px;
      border-bottom: 6px solid var(--gold);
    }
    header h1 { margin: 0; font-size: 24px; letter-spacing: 0; }
    header p { margin: 4px 0 0; color: #dbeafe; }
    main {
      width: min(1180px, calc(100% - 32px));
      margin: 18px auto 28px;
      display: grid;
      grid-template-columns: 330px 1fr;
      gap: 16px;
    }
    section, aside {
      background: var(--white);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    h2 { margin: 0 0 12px; font-size: 17px; }
    label { display: block; font-size: 13px; font-weight: 650; margin: 11px 0 5px; }
    input, select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 9px 10px;
      font: inherit;
      background: #fff;
    }
    .drop {
      border: 2px dashed #9fb8d3;
      background: #f8fbff;
      border-radius: 8px;
      padding: 18px;
      text-align: center;
      cursor: pointer;
    }
    .drop strong { color: var(--blue); }
    .drop input { display: none; }
    .grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
    .actions { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 14px; }
    button {
      border: 0;
      border-radius: 6px;
      padding: 10px 12px;
      font-weight: 700;
      cursor: pointer;
      background: var(--blue);
      color: white;
    }
    button.secondary { background: #e7edf4; color: var(--blue-dark); }
    button.gold { background: var(--gold); color: #1f2933; }
    button.green { background: var(--green); }
    button:disabled { opacity: .55; cursor: not-allowed; }
    .status {
      margin-top: 12px;
      padding: 12px;
      background: var(--soft);
      border: 1px solid var(--line);
      border-radius: 8px;
      min-height: 84px;
      white-space: pre-wrap;
      font-family: Consolas, monospace;
      font-size: 13px;
    }
    .docs { margin-top: 12px; display: grid; gap: 8px; }
    .doc {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: center;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
    }
    .doc a { color: var(--blue); font-weight: 700; text-decoration: none; }
    .pill { color: var(--muted); font-size: 12px; }
    .tabs { display: flex; gap: 8px; margin-bottom: 12px; }
    .tab { background: #e7edf4; color: var(--blue-dark); }
    .tab.active { background: var(--blue-dark); color: white; }
    .panel { display: none; }
    .panel.active { display: block; }
    .row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
    .toolbar { display: flex; gap: 10px; flex-wrap: wrap; align-items: end; margin: 12px 0; }
    .toolbar > div { min-width: 180px; }
    .table-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 8px; margin-top: 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; background: white; }
    th, td { border-bottom: 1px solid var(--line); padding: 8px; text-align: left; white-space: nowrap; }
    th { background: #f4f7fb; color: var(--blue-dark); position: sticky; top: 0; z-index: 1; }
    td input, td select { min-width: 74px; padding: 6px 7px; }
    td.name { min-width: 220px; font-weight: 650; }
    @media (max-width: 900px) {
      main { grid-template-columns: 1fr; }
      .grid, .actions, .row { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <h1>Gestion scolaire</h1>
    <p>Importer un fichier Excel, choisir les feuilles trimestrielles et générer les documents.</p>
  </header>
  <main>
    <aside>
      <h2>1. Fichier Excel</h2>
      <label class="drop" id="drop">
        <strong>Choisir ou déposer un fichier</strong><br>
        <span class="pill">.xlsx ou .xls</span>
        <input id="file" type="file" accept=".xlsx,.xls">
      </label>
      <div id="fileName" class="pill" style="margin-top:8px;">Aucun fichier chargé</div>

      <label>Année scolaire</label>
      <input id="annee" value="ANNEE SCOLAIRE 2025-2026">
      <label>CEB</label>
      <input id="ceb" value="CEB DE NAMISSIGUIMA">
      <label>École</label>
      <input id="ecole" value="ECOLE DE NAMISSIGUIMA B">
      <label>Classe</label>
      <input id="classe" value="CLASSE DE CP2 A">
    </aside>

    <section>
      <div class="tabs">
        <button class="tab active" data-panel="trim">Trimestres</button>
        <button class="tab" data-panel="annual">Annuel</button>
        <button class="tab" data-panel="proposal">Proposition</button>
        <button class="tab" data-panel="syn">Synoptiques</button>
        <button class="tab" data-panel="students">Élèves</button>
        <button class="tab" data-panel="grades">Notes</button>
      </div>

      <div class="grid">
        <div><label>Feuille T1</label><select id="1er Trimestre"></select></div>
        <div><label>Feuille T2</label><select id="2eme Trimestre"></select></div>
        <div><label>Feuille T3</label><select id="3eme Trimestre"></select></div>
      </div>

      <div id="trim" class="panel active">
        <h2>Bulletins trimestriels</h2>
        <div class="actions">
          <button onclick="generate('trimester',' 1er Trimestre')">Bulletins T1</button>
          <button onclick="generate('trimester',' 2eme Trimestre')">Bulletins T2</button>
          <button onclick="generate('trimester',' 3eme Trimestre')">Bulletins T3</button>
          <button class="gold" onclick="generate('all_trimesters')">Générer les 3 trimestres</button>
        </div>
      </div>

      <div id="annual" class="panel">
        <h2>Bulletin annuel</h2>
        <div class="actions">
          <button class="green" onclick="generate('annual')">Générer les bulletins annuels</button>
          <button class="gold" onclick="generate('proposal')">Proposition de fin d'année</button>
        </div>
      </div>

      <div id="proposal" class="panel">
        <h2>Proposition de fin d'année</h2>
        <p class="pill">Génère le document Word avec les feuilles sélectionnées pour les trois trimestres.</p>
        <div class="actions">
          <button class="gold" onclick="generate('proposal')">Générer la proposition</button>
        </div>
      </div>

      <div id="syn" class="panel">
        <h2>Tableaux synoptiques</h2>
        <div class="actions">
          <button class="secondary" onclick="generate('synoptic','1er Trimestre')">Synoptique T1</button>
          <button class="secondary" onclick="generate('synoptic','2eme Trimestre')">Synoptique T2</button>
          <button class="secondary" onclick="generate('synoptic','3eme Trimestre')">Synoptique T3</button>
        </div>
      </div>

      <div id="students" class="panel">
        <h2>Inscription des élèves</h2>
        <div class="actions">
          <button class="secondary" onclick="importToDb()">Importer le fichier dans la base</button>
          <button class="green" onclick="exportDbExcel()">Exporter la base vers Excel</button>
        </div>
        <div class="row">
          <div><label>Classe</label><input id="studentClasse" value="CP2 A"></div>
          <div><label>Nom & Prénom</label><input id="studentName" placeholder="Nom et prénom"></div>
          <div><label>Sexe</label><select id="studentSexe"><option></option><option>F</option><option>M</option></select></div>
          <div><label>Age</label><input id="studentAge" type="number" min="3" max="30"></div>
        </div>
        <div class="row">
          <div><label>Scolarité</label><input id="studentScolarite" type="number" min="0" max="20"></div>
          <div><label>Photo</label><input id="studentPhoto" placeholder="chemin de la photo"></div>
          <div style="display:flex;align-items:end;"><button onclick="saveStudent()">Inscrire / modifier</button></div>
          <div style="display:flex;align-items:end;"><button class="secondary" onclick="loadStudents()">Actualiser</button></div>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Nom & Prénom</th><th>Sexe</th><th>Age</th><th>Scolarité</th><th>Photo</th></tr></thead>
            <tbody id="studentsRows"></tbody>
          </table>
        </div>
      </div>

      <div id="grades" class="panel">
        <h2>Saisie des notes</h2>
        <div class="toolbar">
          <div><label>Classe</label><select id="gradeClasse"></select></div>
          <div><label>Trimestre</label><select id="gradeTrimester"><option>1er Trimestre</option><option>2eme Trimestre</option><option>3eme Trimestre</option></select></div>
          <button onclick="loadGrades()">Charger les notes</button>
          <button class="green" onclick="saveGrades()">Enregistrer les notes</button>
        </div>
        <div class="table-wrap">
          <table id="gradesTable">
            <thead id="gradesHead"></thead>
            <tbody id="gradesRows"></tbody>
          </table>
        </div>
      </div>

      <div id="status" class="status">Prêt.</div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:14px;gap:10px;">
        <h2 style="margin:0;">Documents générés</h2>
        <button class="secondary" onclick="loadDocuments()">Actualiser la liste</button>
      </div>
      <div class="pill" id="outputDir" style="margin-top:6px;"></div>
      <div id="docs" class="docs"></div>
    </section>
  </main>

  <script>
    let currentFile = "";
    let matieres = [];
    const names = {"1er Trimestre": "PREMIER TRIMESTRE", "2eme Trimestre": "DEUXIEME TRIMESTRE", "3eme Trimestre": "TROISIEME TRIMESTRE"};
    const statusBox = document.getElementById("status");
    const docs = document.getElementById("docs");
    const outputDir = document.getElementById("outputDir");
    initConfig();
    initDefaultFile();
    loadClasses();
    loadDocuments();

    document.querySelectorAll(".tab").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tab").forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
        btn.classList.add("active");
        document.getElementById(btn.dataset.panel).classList.add("active");
      });
    });

    document.getElementById("file").addEventListener("change", e => upload(e.target.files[0]));
    document.getElementById("studentClasse").addEventListener("change", loadStudents);
    document.getElementById("gradeClasse").addEventListener("change", loadGrades);
    document.getElementById("gradeTrimester").addEventListener("change", loadGrades);
    document.getElementById("drop").addEventListener("dragover", e => e.preventDefault());
    document.getElementById("drop").addEventListener("drop", e => {
      e.preventDefault();
      upload(e.dataTransfer.files[0]);
    });

    async function upload(file) {
      if (!file) return;
      statusBox.textContent = "Chargement du fichier...";
      const data = new FormData();
      data.append("file", file);
      const res = await fetch("/api/upload", {method: "POST", body: data});
      const json = await res.json();
      if (!res.ok) return showError(json.error);
      currentFile = json.path;
      document.getElementById("fileName").textContent = json.name;
      fillSheets(json.sheets);
      statusBox.textContent = "Feuilles détectées : " + json.sheets.join(", ");
    }

    async function initDefaultFile() {
      const res = await fetch("/api/default");
      const json = await res.json();
      if (!res.ok || !json.path) return;
      currentFile = json.path;
      document.getElementById("fileName").textContent = json.name + " (chargé automatiquement)";
      fillSheets(json.sheets);
      statusBox.textContent = "Fichier chargé automatiquement : " + json.name + "\nFeuilles détectées : " + json.sheets.join(", ");
    }

    async function initConfig() {
      const res = await fetch("/api/config");
      const json = await res.json();
      matieres = json.matieres || [];
    }

    function fillSheets(sheets) {
      ["1er Trimestre","2eme Trimestre","3eme Trimestre"].forEach((id, index) => {
        const select = document.getElementById(id);
        select.innerHTML = "";
        sheets.forEach(s => {
          const opt = document.createElement("option");
          opt.value = s;
          opt.textContent = s;
          select.appendChild(opt);
        });
        if (sheets.length) select.value = sheets[Math.min(index, sheets.length - 1)];
      });
    }

    function payload(kind, term) {
      return {
        kind, term,
        file: currentFile,
        output: "bulletins",
        annee: document.getElementById("annee").value,
        ceb: document.getElementById("ceb").value,
        ecole: document.getElementById("ecole").value,
        classe: document.getElementById("classe").value,
        selections: {
          "1er Trimestre": document.getElementById("1er Trimestre").value,
          "2eme Trimestre": document.getElementById("2eme Trimestre").value,
          "3eme Trimestre": document.getElementById("3eme Trimestre").value
        }
      };
    }

    async function generate(kind, term="") {
      if (!currentFile) return showError("Importe d'abord un fichier Excel.");
      docs.innerHTML = "";
      statusBox.textContent = "Génération en cours...";
      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify(payload(kind, term))
        });
        const json = await res.json();
        if (!res.ok) return showError(json.error);
        statusBox.textContent = json.message + (json.warnings.length ? "\n\nAvertissements:\n- " + json.warnings.join("\n- ") : "");
        await loadDocuments();
      } catch (error) {
        showError("Le serveur ne répond pas : " + error.message);
      }
    }

    function addDoc(doc) {
      const row = document.createElement("div");
      row.className = "doc";
      const detail = doc.pages ? `${doc.pages} page(s), ${doc.students || ""} élève(s)` : "Document disponible";
      row.innerHTML = `<div><strong>${doc.label}</strong><br><span class="pill">${detail}</span></div><a href="${doc.url}" target="_blank">Ouvrir</a>`;
      docs.appendChild(row);
    }

    async function loadDocuments() {
      const res = await fetch("/api/documents");
      const json = await res.json();
      docs.innerHTML = "";
      outputDir.textContent = json.output_dir ? "Dossier : " + json.output_dir : "";
      if (!json.documents.length) {
        docs.innerHTML = '<div class="pill">Aucun document généré pour le moment.</div>';
        return;
      }
      json.documents.forEach(addDoc);
    }

    function showError(message) {
      statusBox.textContent = "Erreur : " + message;
    }

    function classeSimple() {
      return document.getElementById("classe").value.replace("CLASSE DE ", "").trim() || "Classe";
    }

    async function loadClasses() {
      const res = await fetch("/api/classes");
      const json = await res.json();
      const select = document.getElementById("gradeClasse");
      select.innerHTML = "";
      const values = json.classes && json.classes.length ? json.classes : [classeSimple()];
      values.forEach(classe => {
        const opt = document.createElement("option");
        opt.value = classe;
        opt.textContent = classe;
        select.appendChild(opt);
      });
      document.getElementById("studentClasse").value = values[0] || classeSimple();
      await loadStudents();
    }

    async function importToDb() {
      if (!currentFile) return showError("Importe d'abord un fichier Excel.");
      statusBox.textContent = "Import dans la base en cours...";
      const res = await fetch("/api/db/import", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(payload("db_import"))
      });
      const json = await res.json();
      if (!res.ok) return showError(json.error);
      statusBox.textContent = `${json.students} élève(s) et ${json.notes} note(s) importés dans la classe ${json.classe}.`;
      await loadClasses();
    }

    async function exportDbExcel() {
      const classe = document.getElementById("studentClasse").value || document.getElementById("gradeClasse").value || classeSimple();
      statusBox.textContent = "Export Excel depuis la base...";
      const res = await fetch("/api/db/export", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({classe})
      });
      const json = await res.json();
      if (!res.ok) return showError(json.error);
      currentFile = json.path;
      document.getElementById("fileName").textContent = json.name + " (exporté depuis la base)";
      fillSheets(json.sheets);
      statusBox.textContent = "Excel généré depuis la base : " + json.name + "\nVous pouvez maintenant générer les bulletins avec ce fichier.";
    }

    async function saveStudent() {
      const data = {
        classe: document.getElementById("studentClasse").value.trim(),
        nom_prenom: document.getElementById("studentName").value.trim(),
        sexe: document.getElementById("studentSexe").value,
        age: document.getElementById("studentAge").value,
        scolarite: document.getElementById("studentScolarite").value,
        photo: document.getElementById("studentPhoto").value.trim()
      };
      const res = await fetch("/api/students", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(data)
      });
      const json = await res.json();
      if (!res.ok) return showError(json.error);
      statusBox.textContent = "Élève enregistré : " + json.student.nom_prenom;
      document.getElementById("studentName").value = "";
      document.getElementById("studentAge").value = "";
      document.getElementById("studentScolarite").value = "";
      document.getElementById("studentPhoto").value = "";
      await loadClasses();
    }

    async function loadStudents() {
      const classe = document.getElementById("studentClasse").value || classeSimple();
      const res = await fetch("/api/students?classe=" + encodeURIComponent(classe));
      const json = await res.json();
      const tbody = document.getElementById("studentsRows");
      tbody.innerHTML = "";
      (json.students || []).forEach(eleve => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td class="name">${eleve.nom_prenom}</td><td>${eleve.sexe || ""}</td><td>${eleve.age || ""}</td><td>${eleve.scolarite || ""}</td><td>${eleve.photo || ""}</td>`;
        tbody.appendChild(tr);
      });
    }

    async function loadGrades() {
      const classe = document.getElementById("gradeClasse").value || classeSimple();
      const trimestre = document.getElementById("gradeTrimester").value;
      const res = await fetch(`/api/notes?classe=${encodeURIComponent(classe)}&trimestre=${encodeURIComponent(trimestre)}`);
      const json = await res.json();
      if (!res.ok) return showError(json.error);
      const head = document.getElementById("gradesHead");
      const body = document.getElementById("gradesRows");
      head.innerHTML = `<tr><th>Élève</th>${matieres.map(m => `<th>${m.libelle}</th>`).join("")}</tr>`;
      body.innerHTML = "";
      (json.students || []).forEach(eleve => {
        const tr = document.createElement("tr");
        tr.dataset.id = eleve.id;
        tr.innerHTML = `<td class="name">${eleve.nom_prenom}</td>` + matieres.map(m => {
          const value = (eleve.notes && eleve.notes[m.colonne]) || "";
          return `<td><input type="number" step="0.01" min="0" max="10" data-col="${m.colonne}" value="${value}"></td>`;
        }).join("");
        body.appendChild(tr);
      });
      statusBox.textContent = `${(json.students || []).length} élève(s) chargés pour ${classe}, ${trimestre}.`;
    }

    async function saveGrades() {
      const classe = document.getElementById("gradeClasse").value || classeSimple();
      const trimestre = document.getElementById("gradeTrimester").value;
      const rows = [...document.querySelectorAll("#gradesRows tr")].map(tr => {
        const notes = {};
        tr.querySelectorAll("input[data-col]").forEach(input => notes[input.dataset.col] = input.value);
        return {id: Number(tr.dataset.id), notes};
      });
      const res = await fetch("/api/notes", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({classe, trimestre, rows})
      });
      const json = await res.json();
      if (!res.ok) return showError(json.error);
      statusBox.textContent = `${json.saved} note(s) enregistrées pour ${classe}, ${trimestre}.`;
    }
  </script>
</body>
</html>"""


def json_response(handler, payload, status=200):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def safe_path(value):
    path = Path(value).resolve()
    if ROOT not in path.parents and path != ROOT:
        raise bulletins.ValidationErreur("Chemin non autorisé.")
    return path


def documents_existants():
    OUTPUTS.mkdir(exist_ok=True)
    documents = []
    files = list(OUTPUTS.glob("*.pdf")) + list(OUTPUTS.glob("*.docx"))
    for path in sorted(files, key=lambda p: p.stat().st_mtime, reverse=True):
        documents.append({
            "label": path.stem.replace("_", " ").title(),
            "students": "",
            "pages": "",
            "url": "/download?path=" + quote(str(path.resolve())),
        })
    return documents


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/api/documents":
            json_response(self, {"documents": documents_existants(), "output_dir": str(OUTPUTS.resolve())})
            return
        if parsed.path == "/api/default":
            self.default_file()
            return
        if parsed.path == "/api/config":
            json_response(self, {"matieres": bulletins.MATIERES})
            return
        if parsed.path == "/api/classes":
            json_response(self, {"classes": school_db.classes()})
            return
        if parsed.path == "/api/students":
            classe = parse_qs(parsed.query).get("classe", [""])[0]
            json_response(self, {"students": school_db.students(classe)})
            return
        if parsed.path == "/api/notes":
            query = parse_qs(parsed.query)
            classe = query.get("classe", [""])[0]
            trimestre = query.get("trimestre", [""])[0]
            json_response(self, {"students": school_db.notes_for_class(classe, trimestre)})
            return
        if parsed.path == "/download":
            try:
                path = safe_path(parse_qs(parsed.query).get("path", [""])[0])
                if not path.exists():
                    raise FileNotFoundError(path)
                data = path.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", mimetypes.guess_type(path.name)[0] or "application/octet-stream")
                self.send_header("Content-Disposition", f'inline; filename="{path.name}"')
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            except Exception as exc:
                json_response(self, {"error": str(exc)}, 404)
            return
        json_response(self, {"error": "Route introuvable."}, 404)

    def do_POST(self):
        if self.path == "/api/upload":
            self.upload()
            return
        if self.path == "/api/generate":
            self.generate()
            return
        if self.path == "/api/db/import":
            self.import_db()
            return
        if self.path == "/api/db/export":
            self.export_db()
            return
        if self.path == "/api/students":
            self.save_student()
            return
        if self.path == "/api/notes":
            self.save_notes()
            return
        json_response(self, {"error": "Route introuvable."}, 404)

    def upload(self):
        try:
            UPLOADS.mkdir(exist_ok=True)
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
            item = form["file"]
            filename = Path(item.filename).name
            target = UPLOADS / filename
            target.write_bytes(item.file.read())
            sheets = bulletins.lister_feuilles(target)
            json_response(self, {"name": filename, "path": str(target), "sheets": sheets})
        except Exception as exc:
            json_response(self, {"error": str(exc)}, 400)

    def default_file(self):
        try:
            candidates = sorted(
                list(ROOT.glob("*.xlsx")) + list(ROOT.glob("*.xls")),
                key=lambda path: path.stat().st_mtime,
                reverse=True,
            )
            candidates.extend([ROOT / "Notes.xlsx", ROOT / "Notes1.xlsx"])
            target = next((path for path in candidates if path.exists()), None)
            if target is None:
                json_response(self, {"path": "", "sheets": [], "name": ""})
                return
            sheets = bulletins.lister_feuilles(target)
            json_response(self, {"name": target.name, "path": str(target), "sheets": sheets})
        except Exception as exc:
            json_response(self, {"error": str(exc)}, 400)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def import_db(self):
        try:
            data = self.read_json()
            file_path = safe_path(data["file"])
            classe = data.get("classe", bulletins.CLASSE)
            selections = data.get("selections", {})
            result = school_db.import_excel(file_path, selections, classe)
            json_response(self, result)
        except Exception as exc:
            traceback.print_exc()
            json_response(self, {"error": str(exc)}, 400)

    def export_db(self):
        try:
            data = self.read_json()
            output = school_db.export_excel(data.get("classe", "Classe"))
            sheets = bulletins.lister_feuilles(output)
            json_response(self, {"name": output.name, "path": str(output.resolve()), "sheets": sheets})
        except Exception as exc:
            traceback.print_exc()
            json_response(self, {"error": str(exc)}, 400)

    def save_student(self):
        try:
            student = school_db.save_student(self.read_json())
            json_response(self, {"student": student})
        except Exception as exc:
            traceback.print_exc()
            json_response(self, {"error": str(exc)}, 400)

    def save_notes(self):
        try:
            data = self.read_json()
            result = school_db.save_notes(data.get("classe", ""), data.get("trimestre", ""), data.get("rows", []))
            json_response(self, result)
        except Exception as exc:
            traceback.print_exc()
            json_response(self, {"error": str(exc)}, 400)

    def generate(self):
        try:
            data = self.read_json()
            file_path = safe_path(data["file"])
            output = OUTPUTS
            annee = data.get("annee", bulletins.ANNEE)
            ceb = data.get("ceb", bulletins.CEB)
            ecole = data.get("ecole", bulletins.ECOLE)
            classe = data.get("classe", bulletins.CLASSE)
            selections = data.get("selections", {})
            docs = []
            warnings = []

            def add(label, result):
                warnings.extend(result.get("avertissements", []))
                path = Path(result.get("pdf_path") or result.get("docx_path")).resolve()
                docs.append({
                    "label": label,
                    "students": result["eleves"],
                    "pages": result["pages"],
                    "url": "/download?path=" + quote(str(path)),
                })

            kind = data.get("kind")
            term = data.get("term")
            if kind == "trimester":
                label = {"1er Trimestre": "PREMIER TRIMESTRE", "2eme Trimestre": "DEUXIEME TRIMESTRE", "3eme Trimestre": "TROISIEME TRIMESTRE"}[term]
                add("Bulletins " + label, bulletins.generer_bulletins_trimestre(file_path, selections[term], label, output, annee, ceb, ecole, classe))
            elif kind == "all_trimesters":
                for code, label in [("1er Trimestre", "PREMIER TRIMESTRE"), ("2eme Trimestre", "DEUXIEME TRIMESTRE"), ("3eme Trimestre", "TROISIEME TRIMESTRE")]:
                    add("Bulletins " + label, bulletins.generer_bulletins_trimestre(file_path, selections[code], label, output, annee, ceb, ecole, classe))
            elif kind == "annual":
                selected = [(code, selections[code]) for code in ["1er Trimestre", "2eme Trimestre", "3eme Trimestre"]]
                add("Bulletins annuels", bulletins.generer_bulletins_annuels(file_path, selected, output, annee, ceb, ecole, classe))
            elif kind == "proposal":
                selected = [(code, selections[code]) for code in ["1er Trimestre", "2eme Trimestre", "3eme Trimestre"]]
                add("Proposition de fin d'année", bulletins.generer_proposition_fin_annee(file_path, selected, output, annee=annee.replace("ANNEE SCOLAIRE ", ""), ceb=ceb.replace("CEB DE ", ""), ecole=ecole.replace("ECOLE DE ", ""), classe=classe))
            elif kind == "synoptic":
                label = {"1er Trimestre": "PREMIER TRIMESTRE", "2eme Trimestre": "DEUXIEME TRIMESTRE", "3eme Trimestre": "TROISIEME TRIMESTRE"}[term]
                add("Tableau synoptique " + label, bulletins.generer_tableau_synoptique(file_path, selections[term], label, output, annee, ceb, ecole, classe))
            else:
                raise bulletins.ValidationErreur("Action inconnue.")

            json_response(self, {"message": "Génération terminée.", "documents": docs, "warnings": warnings})
        except Exception as exc:
            traceback.print_exc()
            json_response(self, {"error": str(exc)}, 400)


def run(host="127.0.0.1", port=8013):
    print(f"Interface disponible sur http://{host}:{port}")
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == "__main__":
    run()
