from pathlib import Path
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import generer_bulletins as bulletins


TRIMESTRES = [("Moyenne du 1er Trimestre", "PREMIER TRIMESTRE"), ("Moyenne du 2e Trimestre", "DEUXIEME TRIMESTRE"), ("Moyenne du 3e Trimestre", "TROISIEME TRIMESTRE")]


class InterfaceBulletins(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestion scolaire - Bulletins")
        self.geometry("940x720")
        self.minsize(860, 640)
        self.pdf_genere = None
        self.feuilles = []
        self.variables = {
            "fichier_excel": tk.StringVar(value=str(Path(bulletins.FICHIER_EXCEL).resolve())),
            "dossier_sortie": tk.StringVar(value=str(Path(bulletins.DOSSIER_SORTIE).resolve())),
            "annee": tk.StringVar(value=bulletins.ANNEE),
            "ceb": tk.StringVar(value=bulletins.CEB),
            "ecole": tk.StringVar(value=bulletins.ECOLE),
            "classe": tk.StringVar(value=bulletins.CLASSE),
            "T1": tk.StringVar(),
            "T2": tk.StringVar(),
            "T3": tk.StringVar(),
        }
        self.combos = {code: [] for code, _ in TRIMESTRES}
        self._style()
        self._build()
        self.charger_feuilles_depuis_fichier(silencieux=True)

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f6f7f9")
        style.configure("TLabelframe", background="#f6f7f9")
        style.configure("TLabelframe.Label", background="#f6f7f9", font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background="#f6f7f9", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))

    def _build(self):
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)
        ttk.Label(root, text="Gestion scolaire", font=("Segoe UI", 17, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
        self._section_parametres(root).grid(row=1, column=0, sticky="ew", pady=(0, 10))
        notebook = ttk.Notebook(root)
        notebook.grid(row=2, column=0, sticky="nsew")
        notebook.add(self._onglet_trimestres(notebook), text="Trimestres")
        notebook.add(self._onglet_annuel(notebook), text="Annuel")
        notebook.add(self._onglet_synoptiques(notebook), text="Synoptiques")
        self._section_journal(root).grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        self._section_actions(root).grid(row=4, column=0, sticky="ew", pady=(10, 0))

    def _section_parametres(self, parent):
        frame = ttk.LabelFrame(parent, text="Parametres", padding=10)
        for col in (1, 3):
            frame.columnconfigure(col, weight=1)
        self._entry(frame, 0, 0, "Fichier Excel", "fichier_excel", self.choisir_excel)
        self._entry(frame, 1, 0, "Dossier sortie", "dossier_sortie", self.choisir_dossier)
        self._entry(frame, 2, 0, "Annee", "annee")
        self._entry(frame, 0, 2, "CEB", "ceb")
        self._entry(frame, 1, 2, "Ecole", "ecole")
        self._entry(frame, 2, 2, "Classe", "classe")
        ttk.Button(frame, text="Charger les feuilles", command=self.charger_feuilles_depuis_fichier).grid(row=3, column=0, columnspan=4, sticky="e", pady=(8, 0))
        return frame

    def _entry(self, parent, row, col, label, key, command=None):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=(0, 6), pady=3)
        ttk.Entry(parent, textvariable=self.variables[key]).grid(row=row, column=col + 1, sticky="ew", pady=3, padx=(0, 8))
        if command:
            ttk.Button(parent, text="Parcourir", command=command).grid(row=row, column=col + 2, sticky="w", pady=3, padx=(0, 8))

    def _selecteurs_feuilles(self, parent):
        frame = ttk.LabelFrame(parent, text="Selection manuelle des feuilles", padding=10)
        for index, (code, libelle) in enumerate(TRIMESTRES):
            ttk.Label(frame, text=libelle).grid(row=index, column=0, sticky="w", padx=(0, 8), pady=5)
            combo = ttk.Combobox(frame, textvariable=self.variables[code], values=self.feuilles, state="readonly", width=32)
            combo.grid(row=index, column=1, sticky="ew", pady=5)
            self.combos[code].append(combo)
        frame.columnconfigure(1, weight=1)
        return frame

    def _onglet_trimestres(self, parent):
        frame = ttk.Frame(parent, padding=12)
        frame.columnconfigure(0, weight=1)
        self._selecteurs_feuilles(frame).grid(row=0, column=0, sticky="ew", pady=(0, 12))
        actions = ttk.LabelFrame(frame, text="Bulletins trimestriels", padding=10)
        actions.grid(row=1, column=0, sticky="ew")
        for index, (code, libelle) in enumerate(TRIMESTRES):
            ttk.Button(actions, text=f"Generer {libelle}", command=lambda c=code, l=libelle: self.generer_trimestre(c, l)).grid(row=0, column=index, padx=5, pady=5, sticky="ew")
            actions.columnconfigure(index, weight=1)
        ttk.Button(actions, text="Generer les 3 trimestres", command=self.generer_tous_trimestres, style="Primary.TButton").grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        return frame

    def _onglet_annuel(self, parent):
        frame = ttk.Frame(parent, padding=12)
        frame.columnconfigure(0, weight=1)
        self._selecteurs_feuilles(frame).grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Button(frame, text="Generer les bulletins annuels", command=self.generer_annuel, style="Primary.TButton").grid(row=1, column=0, sticky="ew")
        return frame

    def _onglet_synoptiques(self, parent):
        frame = ttk.Frame(parent, padding=12)
        frame.columnconfigure(0, weight=1)
        self._selecteurs_feuilles(frame).grid(row=0, column=0, sticky="ew", pady=(0, 12))
        actions = ttk.LabelFrame(frame, text="Tableaux synoptiques", padding=10)
        actions.grid(row=1, column=0, sticky="ew")
        for index, (code, libelle) in enumerate(TRIMESTRES):
            ttk.Button(actions, text=f"Synoptique {libelle}", command=lambda c=code, l=libelle: self.generer_synoptique(c, l)).grid(row=0, column=index, padx=5, pady=5, sticky="ew")
            actions.columnconfigure(index, weight=1)
        return frame

    def _section_journal(self, parent):
        frame = ttk.LabelFrame(parent, text="Journal", padding=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.journal = ScrolledText(frame, height=9, wrap="word", font=("Consolas", 9))
        self.journal.grid(row=0, column=0, sticky="nsew")
        return frame

    def _section_actions(self, parent):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        self.bouton_pdf = ttk.Button(frame, text="Ouvrir le dernier PDF", command=self.ouvrir_pdf, state="disabled")
        self.bouton_pdf.grid(row=0, column=1, padx=(0, 8))
        ttk.Button(frame, text="Quitter", command=self.destroy).grid(row=0, column=2)
        return frame

    def choisir_excel(self):
        chemin = filedialog.askopenfilename(title="Choisir le fichier Excel", filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")])
        if chemin:
            self.variables["fichier_excel"].set(chemin)
            self.charger_feuilles_depuis_fichier()

    def choisir_dossier(self):
        chemin = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if chemin:
            self.variables["dossier_sortie"].set(chemin)

    def charger_feuilles_depuis_fichier(self, silencieux=False):
        try:
            self.feuilles = bulletins.lister_feuilles(self.variables["fichier_excel"].get().strip())
        except Exception as exc:
            if not silencieux:
                messagebox.showerror("Chargement impossible", str(exc))
            return
        for index, (code, _) in enumerate(TRIMESTRES):
            for combo in self.combos.get(code, []):
                combo["values"] = self.feuilles
            if not self.variables[code].get() and self.feuilles:
                self.variables[code].set(self.feuilles[min(index, len(self.feuilles) - 1)])
        if not silencieux:
            self.log(f"Feuilles chargees : {', '.join(self.feuilles)}")

    def selections(self):
        result = []
        for code, libelle in TRIMESTRES:
            feuille = self.variables[code].get().strip()
            if not feuille:
                raise bulletins.ValidationErreur(f"Aucune feuille selectionnee pour {libelle}.")
            result.append((code, feuille))
        return result

    def params(self):
        return {
            "fichier_excel": self.variables["fichier_excel"].get().strip(),
            "dossier_sortie": self.variables["dossier_sortie"].get().strip(),
            "annee": self.variables["annee"].get().strip(),
            "ceb": self.variables["ceb"].get().strip(),
            "ecole": self.variables["ecole"].get().strip(),
            "classe": self.variables["classe"].get().strip(),
        }

    def log(self, message):
        self.journal.insert("end", message + "\n")
        self.journal.see("end")

    def afficher_resultat(self, resultat):
        self.pdf_genere = Path(resultat["pdf_path"]).resolve()
        self.bouton_pdf.configure(state="normal")
        self.log(f"PDF : {self.pdf_genere}")
        self.log(f"Eleves : {resultat['eleves']} | Pages : {resultat['pages']}")
        if resultat.get("avertissements"):
            self.log("Avertissements :")
            for avertissement in resultat["avertissements"]:
                self.log(f"- {avertissement}")

    def executer(self, titre, action):
        self.journal.delete("1.0", "end")
        self.log(titre)
        self.update_idletasks()
        try:
            resultat = action()
        except bulletins.ValidationErreur as exc:
            self.log(str(exc))
            messagebox.showerror("Generation arretee", str(exc))
        except Exception as exc:
            self.log(str(exc))
            messagebox.showerror("Erreur", str(exc))
        else:
            self.afficher_resultat(resultat)
            messagebox.showinfo("Termine", f"Document genere :\n{self.pdf_genere}")

    def generer_trimestre(self, code, libelle):
        feuille = self.variables[code].get().strip()
        p = self.params()
        self.executer(f"Generation {libelle}...", lambda: bulletins.generer_bulletins_trimestre(p["fichier_excel"], feuille, libelle, p["dossier_sortie"], p["annee"], p["ceb"], p["ecole"], p["classe"]))

    def generer_tous_trimestres(self):
        p = self.params()
        def action():
            dernier = None
            total = 0
            for code, libelle in TRIMESTRES:
                dernier = bulletins.generer_bulletins_trimestre(p["fichier_excel"], self.variables[code].get().strip(), libelle, p["dossier_sortie"], p["annee"], p["ceb"], p["ecole"], p["classe"])
                total += dernier["pages"]
                self.log(f"{libelle} genere.")
            dernier["pages"] = total
            return dernier
        self.executer("Generation des trois trimestres...", action)

    def generer_annuel(self):
        p = self.params()
        self.executer("Generation des bulletins annuels...", lambda: bulletins.generer_bulletins_annuels(p["fichier_excel"], self.selections(), p["dossier_sortie"], p["annee"], p["ceb"], p["ecole"], p["classe"]))

    def generer_synoptique(self, code, libelle):
        p = self.params()
        feuille = self.variables[code].get().strip()
        self.executer(f"Generation tableau synoptique {libelle}...", lambda: bulletins.generer_tableau_synoptique(p["fichier_excel"], feuille, libelle, p["dossier_sortie"], p["annee"], p["ceb"], p["ecole"], p["classe"]))

    def ouvrir_pdf(self):
        if self.pdf_genere and self.pdf_genere.exists():
            os.startfile(self.pdf_genere)
        else:
            messagebox.showwarning("PDF introuvable", "Aucun PDF genere n'est disponible.")


if __name__ == "__main__":
    app = InterfaceBulletins()
    app.mainloop()
