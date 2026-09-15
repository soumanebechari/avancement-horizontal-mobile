# -*- coding: utf-8 -*-
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.metrics import dp
from kivy.graphics import Color, RoundedRectangle
from kivy.core.window import Window
import os, csv, datetime, math, json, shutil

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

APP_DIR = "/storage/emulated/0/AvancementHorizontal"
HISTORY_FILE = os.path.join(APP_DIR, "historique_promotions.csv")
RESULT_FILE = os.path.join(APP_DIR, "Resultat_Avancement.xlsx")
RESULTS_HISTORY_FILE = os.path.join(APP_DIR, "historique_resultats_promotion.csv")
DATABASE_FILE = os.path.join(APP_DIR, "base_de_donnees.json")
BACKUP_PREFIX = "sauvegarde_avancement_"

CATEGORIES = ("Toutes", "Emploi supérieur", "Cadre Supérieur", "Cadre", "Maîtrise", "Exécution")
SANCTIONS = ("Aucune", "2ème degré", "3ème degré")
NOTE_YEARS = range(2022, 2031)
EMPLOYEE_FIELDS = ["Matricule", "Nom et Prénom", "Catégorie", "Date de naissance", "Date de recrutement", "Dernière promotion", "Sanction", "Année de la sanction", "Historique des sanctions"]
HISTORY_FIELDS = ["Matricule", "Nom et Prénom", "Année", "Date d'effet", "Échelon", "Rythme", "Observation"]
RESULT_HISTORY_FIELDS = ["Année de calcul", "Matricule", "Nom et Prénom", "Catégorie", "Dernière promotion", "Moyenne", "Ancienneté", "Rythme proposé", "Prochaine promotion", "Sanction"]
NON_BENEF_FIELDS = ["Année de promotion", "Matricule", "Nom et Prénom", "Catégorie", "Dernière promotion", "Moyenne", "Ancienneté", "Sanction", "Année de la sanction", "Motif de non-bénéfice"]


def norm(v):
    return str(v or "").strip().lower().replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("û", "u")


def format_date(v):
    """Return dates without time, preferably in DD/MM/YYYY format."""
    if v is None or v == "":
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, datetime.date):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    if not s:
        return ""
    # Remove a time part from common ISO/datetime strings.
    if "T" in s:
        s = s.split("T", 1)[0]
    if " " in s and any(ch.isdigit() for ch in s):
        first = s.split(" ", 1)[0]
        if first.count("-") == 2 or first.count("/") == 2:
            s = first
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    # A year-only value is kept as a year; this is useful for promotion years.
    if len(s) == 4 and s.isdigit():
        return s
    return s


def normalize_row_dates(row):
    for key in ("Date de naissance", "Date de recrutement", "Dernière promotion"):
        if key in row:
            row[key] = format_date(row.get(key))
    return row


def date_year(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.year
    s = str(v or "").strip()
    if not s:
        return None
    for part in (s[:4], s[-4:]):
        try:
            y = int(part)
            if 1900 <= y <= 2200:
                return y
        except Exception:
            pass
    return None


def all_note_years(rows=None):
    """Return all notation years, including years added manually by the user."""
    years = set(NOTE_YEARS)
    for row in (rows or []):
        for key in row.keys():
            k = str(key)
            if k.startswith("note_"):
                try:
                    y = int(k[5:])
                    if 1900 <= y <= 2200:
                        years.add(y)
                except Exception:
                    pass
    return sorted(years)


class Card(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.padding = dp(10)
        self.spacing = dp(6)
        with self.canvas.before:
            Color(0.97, 0.98, 0.97, 1)
            self.bg = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(7)])
        self.bind(pos=self._sync, size=self._sync)

    def _sync(self, *_):
        self.bg.pos = self.pos
        self.bg.size = self.size


class MainApp(App):
    def build(self):
        os.makedirs(APP_DIR, exist_ok=True)
        self.file_path = ""
        self.rows = []
        self.last_results = []
        self.last_non_beneficiaries = []
        self.current_year = datetime.datetime.now().year
        self.load_database_safely()
        Window.clearcolor = (0.94, 0.95, 0.94, 1)
        root = BoxLayout(orientation="vertical")
        root.add_widget(self.header())
        root.add_widget(self.navbar())
        scroll = ScrollView(do_scroll_x=False)
        self.main = BoxLayout(orientation="vertical", spacing=dp(9), padding=dp(10), size_hint_y=None)
        self.main.bind(minimum_height=self.main.setter("height"))
        scroll.add_widget(self.main)
        root.add_widget(scroll)
        self.show_home()
        return root

    def header(self):
        h = BoxLayout(size_hint_y=None, height=dp(52), padding=[dp(12), dp(5)])
        with h.canvas.before:
            Color(0.10, 0.48, 0.14, 1)
            self.hbg = RoundedRectangle(pos=h.pos, size=h.size)
        h.bind(pos=lambda *_: setattr(self.hbg, 'pos', h.pos), size=lambda *_: setattr(self.hbg, 'size', h.size))
        title = Label(text="Programme de Calcul de Promotion", bold=True, color=(1, 1, 1, 1), font_size="16sp", halign="left")
        title.bind(size=lambda inst, val: setattr(inst, 'text_size', val))
        h.add_widget(title)
        return h

    def navbar(self):
        bar = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(4), padding=dp(4))
        for text, fn in [("Accueil", self.show_home), ("Employés", self.show_employees), ("Paramètres", self.calcul), ("Historique", self.history)]:
            b = Button(text=text, background_normal="", background_color=(0.86, 0.91, 0.86, 1), color=(0.12, 0.35, 0.15, 1), font_size="12sp")
            b.bind(on_release=fn)
            bar.add_widget(b)
        return bar

    def clear(self):
        self.main.clear_widgets()

    def label(self, text, size=14, bold=False):
        return Label(text=str(text), font_size=f"{size}sp", bold=bold, color=(0.16, 0.24, 0.17, 1), halign="left", valign="middle")

    def btn(self, text, fn, kind="normal", height=44):
        colors = {"green": (0.18, 0.60, 0.25, 1), "orange": (0.95, 0.57, 0.05, 1), "red": (0.78, 0.14, 0.14, 1), "gray": (0.78, 0.79, 0.79, 1), "normal": (0.88, 0.90, 0.88, 1)}
        b = Button(text=text, size_hint_y=None, height=dp(height), background_normal="", background_color=colors[kind], color=(1, 1, 1, 1) if kind in ("green", "orange", "red") else (0.15, 0.18, 0.15, 1), font_size="12sp")
        b.bind(on_release=fn)
        return b

    def show_home(self, *_):
        self.clear()
        card = Card(orientation="vertical", size_hint_y=None, height=dp(205))
        card.add_widget(self.label("Saisie des données des employés", 16, True))
        self.status = self.label(("Fichier : " + os.path.basename(self.file_path) + " | Agents : " + str(len(self.rows))) if self.file_path else "Aucun fichier Excel chargé.", 13)
        card.add_widget(self.status)
        r1 = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(6))
        r1.add_widget(self.btn("Ajouter employé", self.add_employee, "green"))
        r1.add_widget(self.btn("Importer Excel", self.import_excel, "normal"))
        card.add_widget(r1)
        r2 = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(6))
        r2.add_widget(self.btn("Modifier sélection", self.choose_employee, "orange"))
        r2.add_widget(self.btn("Calculer les promotions", self.calcul, "green"))
        card.add_widget(r2)
        self.main.add_widget(card)

        card2 = Card(orientation="vertical", size_hint_y=None, height=dp(170))
        card2.add_widget(self.label("Gestion", 16, True))
        r3 = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(6))
        r3.add_widget(self.btn("Gérer les notations", self.notations, "normal"))
        r3.add_widget(self.btn("Historique promotions", self.history, "normal"))
        card2.add_widget(r3)
        r4 = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(6))
        r4.add_widget(self.btn("Exporter résultat", self.export_last, "orange"))
        r4.add_widget(self.btn("Vider la liste", self.clear_list, "red"))
        card2.add_widget(r4)
        r5 = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(6))
        r5.add_widget(self.btn("Sauvegarder", self.save_backup, "orange"))
        r5.add_widget(self.btn("Restaurer", self.restore_backup, "normal"))
        card2.add_widget(r5)
        r6 = BoxLayout(size_hint_y=None, height=dp(46), spacing=dp(6))
        r6.add_widget(self.btn("Résultats enregistrés", self.saved_results_history, "normal"))
        card2.add_widget(r6)
        card2.add_widget(self.label("Base locale JSON : les données sont conservées sans heure dans les dates. Sauvegarder crée une copie JSON récupérable plus tard. Règle : ancienneté minimale de 3 ans depuis la dernière promotion. Rapide +3 ans, Moyen +4 ans, Lent +5 ans.", 11))
        self.main.add_widget(card2)

        self.quota_card = Card(orientation="vertical", size_hint_y=None, height=dp(130))
        self.main.add_widget(self.quota_card)
        self.refresh_quota()
        self.employee_list_card()

    def refresh_quota(self, total=None):
        if not hasattr(self, 'quota_card'):
            return
        self.quota_card.clear_widgets()
        n = len(self.last_results) if total is None else total
        fast = math.ceil(n * 0.50)
        medium = math.floor(n * 0.30)
        slow = n - fast - medium
        if slow < 0:
            slow = 0
            medium = max(0, n - fast)
        self.quota_card.add_widget(self.label("Limites autorisées de promotion", 15, True))
        self.quota_card.add_widget(self.label(f"Rapide 50% : {fast} place(s)    |    Moyen 30% : {medium} place(s)    |    Lent 20% : {slow} place(s)", 12))
        self.quota_card.add_widget(self.label(f"Candidats éligibles retenus : {n}    |    Places utilisées : {n}", 12, True))
        self.quota_card.add_widget(self.label("Aucun dépassement de quota : la liste finale est limitée aux places calculées.", 11))

    def employee_list_card(self):
        card = Card(orientation="vertical", size_hint_y=None, height=dp(430))
        card.add_widget(self.label("Liste de tous les employés", 16, True))
        if not self.rows:
            card.add_widget(self.label("Aucun employé. Importez un fichier Excel ou ajoutez un employé.", 12))
        else:
            info = self.label("Touchez une ligne pour ouvrir tous les champs et modifier l'employé.", 11)
            card.add_widget(info)
            scroll = ScrollView(do_scroll_x=True, do_scroll_y=True, size_hint_y=None, height=dp(350))
            grid = GridLayout(cols=8, size_hint_y=None, row_default_height=dp(42), spacing=dp(1), padding=dp(1))
            grid.bind(minimum_height=grid.setter('height'))
            heads = ["Matricule", "Nom et Prénom", "Catégorie", "Naissance", "Recrutement", "Dernière promotion", "Sanction", "Modifier"]
            for x in heads:
                grid.add_widget(self.label(x, 9, True))
            for r in self.rows:
                values = [r.get("Matricule", ""), r.get("Nom et Prénom", ""), r.get("Catégorie", ""), r.get("Date de naissance", ""), r.get("Date de recrutement", ""), r.get("Dernière promotion", ""), r.get("Sanction", "")]
                for x in values:
                    grid.add_widget(self.label(x, 9))
                b = self.btn("Ouvrir", lambda inst, row=r: self.edit_employee(row), "green", 38)
                grid.add_widget(b)
            scroll.add_widget(grid)
            card.add_widget(scroll)
        self.main.add_widget(card)

    def show_employees(self, *_):
        self.clear()
        title = Card(orientation="vertical", size_hint_y=None, height=dp(132))
        title.add_widget(self.label("Employés", 18, True))
        title.add_widget(self.label(f"Nombre total : {len(self.rows)}", 12))
        search = TextInput(hint_text="Rechercher par nom ou matricule...", multiline=False, size_hint_y=None, height=dp(42))
        title.add_widget(search)
        self.main.add_widget(title)

        card = Card(orientation="vertical", size_hint_y=None, height=dp(470))
        card.add_widget(self.label("Liste des employés", 16, True))
        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True, size_hint_y=None, height=dp(390))
        grid = GridLayout(cols=8, size_hint_y=None, row_default_height=dp(42), spacing=dp(1), padding=dp(1))
        grid.bind(minimum_height=grid.setter('height'))
        scroll.add_widget(grid); card.add_widget(scroll)
        self.main.add_widget(card)

        def rebuild(*_):
            grid.clear_widgets()
            heads = ["Matricule", "Nom et Prénom", "Catégorie", "Naissance", "Recrutement", "Dernière promotion", "Sanction", "Modifier"]
            for h in heads: grid.add_widget(self.label(h, 9, True))
            q = norm(search.text)
            count = 0
            for r in self.rows:
                txt = norm(str(r.get("Matricule", "")) + " " + str(r.get("Nom et Prénom", "")))
                if q and q not in txt:
                    continue
                count += 1
                values = [r.get("Matricule", ""), r.get("Nom et Prénom", ""), r.get("Catégorie", ""), r.get("Date de naissance", ""), r.get("Date de recrutement", ""), r.get("Dernière promotion", ""), r.get("Sanction", "")]
                for v in values: grid.add_widget(self.label(v, 9))
                grid.add_widget(self.btn("Ouvrir", lambda inst, row=r: self.edit_employee(row), "green", 38))
            if count == 0:
                grid.add_widget(self.label("Aucun résultat.", 12, True))
                for _ in range(7): grid.add_widget(Label())
        search.bind(text=rebuild)
        rebuild()

    def popup(self, title, widget, size=(0.96, 0.90)):
        p = Popup(title=title, content=widget, size_hint=size)
        p.open()
        return p

    def import_excel(self, *_):
        if load_workbook is None:
            self.popup("Erreur", self.label("openpyxl n'est pas installé."))
            return
        box = BoxLayout(orientation="vertical")
        chooser = __import__('kivy.uix.filechooser', fromlist=['FileChooserListView']).FileChooserListView(path="/storage/emulated/0", filters=["*.xlsx", "*.xlsm"])
        box.add_widget(chooser)
        bar = BoxLayout(size_hint_y=None, height=dp(50))
        load = self.btn("Charger", lambda *_: None, "green")
        cancel = self.btn("Annuler", lambda *_: None, "gray")
        bar.add_widget(load); bar.add_widget(cancel); box.add_widget(bar)
        p = self.popup("Choisir le fichier Excel", box)
        def do(*_):
            if not chooser.selection:
                return
            try:
                self.load_excel(chooser.selection[0]); p.dismiss(); self.show_home()
            except Exception as e:
                self.popup("Erreur", self.label(str(e)))
        load.bind(on_release=do); cancel.bind(on_release=lambda *_: p.dismiss())

    def load_excel(self, path):
        wb = load_workbook(path, data_only=False)
        ws = wb.active
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        rows = []
        for values in ws.iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in values):
                continue
            row = {}
            for i, h in enumerate(headers):
                if h:
                    row[h] = values[i] if i < len(values) else None
            normalize_row_dates(row)
            rows.append(row)
        self.file_path = path
        self.rows = rows
        self.last_results = []
        self.save_database()

    def field_widget(self, name, value):
        if name == "Catégorie":
            return Spinner(text=str(value or "Exécution"), values=CATEGORIES[1:], size_hint_y=None, height=dp(42))
        if name == "Sanction":
            return Spinner(text=str(value or "Aucune"), values=SANCTIONS, size_hint_y=None, height=dp(42))
        return TextInput(text="" if value is None else str(value), multiline=False, size_hint_y=None, height=dp(42))

    def employee_form(self, row, title, on_save, allow_delete=False):
        outer = BoxLayout(orientation="vertical")
        scroll = ScrollView(do_scroll_x=False)
        form = GridLayout(cols=2, spacing=dp(5), padding=dp(8), size_hint_y=None)
        form.bind(minimum_height=form.setter('height'))
        fields = {}
        for name in EMPLOYEE_FIELDS:
            form.add_widget(self.label(name, 11, True))
            w = self.field_widget(name, row.get(name, ""))
            fields[name] = w
            form.add_widget(w)
        # Show every notation year, including years added manually in Gérer les notations.
        for year in all_note_years(self.rows):
            key = "note_" + str(year)
            form.add_widget(self.label("Note " + str(year), 11, True))
            f = TextInput(text="" if row.get(key) is None else str(row.get(key)), multiline=False, input_filter='float', size_hint_y=None, height=dp(42))
            fields[key] = f; form.add_widget(f)
        scroll.add_widget(form); outer.add_widget(scroll)
        bar = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(5))
        save = self.btn("Enregistrer", lambda *_: None, "green")
        cancel = self.btn("Fermer", lambda *_: None, "gray")
        bar.add_widget(save); bar.add_widget(cancel)
        if allow_delete:
            delete = self.btn("Supprimer", lambda *_: None, "red")
            bar.add_widget(delete)
        outer.add_widget(bar)
        p = self.popup(title, outer, (0.98, 0.94))
        cancel.bind(on_release=lambda *_: p.dismiss())
        def save_it(*_):
            for name in EMPLOYEE_FIELDS:
                row[name] = fields[name].text
            normalize_row_dates(row)
            for year in NOTE_YEARS:
                key = "note_" + str(year)
                if key in fields:
                    row[key] = fields[key].text
            on_save()
            p.dismiss()
            self.show_home()
        save.bind(on_release=save_it)
        if allow_delete:
            def delete_it(*_):
                # Demander confirmation avant toute suppression définitive.
                box = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(12))
                box.add_widget(self.label("Voulez-vous vraiment supprimer cet employé ?", 13, True))
                box.add_widget(self.label(str(row.get("Nom et Prénom", "")), 12))
                actions = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(8))
                yes = self.btn("Oui, supprimer", lambda *_: None, "red")
                no = self.btn("Annuler", lambda *_: None, "gray")
                actions.add_widget(yes); actions.add_widget(no); box.add_widget(actions)
                confirm = self.popup("Confirmation de suppression", box, (0.88, 0.38))
                no.bind(on_release=lambda *_: confirm.dismiss())
                def confirm_delete(*_):
                    if row in self.rows:
                        self.rows.remove(row)
                    self.save_database()
                    try:
                        if self.file_path: self.save_excel()
                    except Exception:
                        pass
                    confirm.dismiss(); p.dismiss(); self.show_home()
                yes.bind(on_release=confirm_delete)
            delete.bind(on_release=delete_it)
        return p

    def add_employee(self, *_):
        row = {k: "" for k in EMPLOYEE_FIELDS}
        row["Catégorie"] = "Exécution"; row["Sanction"] = "Aucune"
        for y in NOTE_YEARS: row["note_" + str(y)] = ""
        def save():
            self.rows.append(row)
            self.save_database()
            try:
                if self.file_path: self.save_excel()
            except Exception:
                pass
        self.employee_form(row, "Ajouter employé", save, False)

    def choose_employee(self, *_):
        if not self.rows:
            self.popup("Information", self.label("Aucun employé à modifier.")); return
        outer = BoxLayout(orientation="vertical", spacing=dp(5), padding=dp(5))
        search = TextInput(hint_text="Rechercher par nom ou matricule...", multiline=False, size_hint_y=None, height=dp(42))
        outer.add_widget(search)
        scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
        body = GridLayout(cols=1, spacing=dp(3), size_hint_y=None, padding=dp(3)); body.bind(minimum_height=body.setter('height'))
        scroll.add_widget(body); outer.add_widget(scroll)
        close = self.btn("Fermer", lambda *_: None, "gray", 48); outer.add_widget(close)
        p = self.popup("Modifier sélection", outer, (0.97, 0.90)); close.bind(on_release=lambda *_: p.dismiss())

        def rebuild(*_):
            body.clear_widgets()
            q = norm(search.text)
            count = 0
            for r in self.rows:
                text = norm(str(r.get("Matricule", "")) + " " + str(r.get("Nom et Prénom", "")))
                if q and q not in text:
                    continue
                count += 1
                caption = f"{r.get('Matricule','')}  |  {r.get('Nom et Prénom','')}  |  {r.get('Catégorie','')}"
                body.add_widget(self.btn(caption, lambda inst, row=r: self.edit_employee(row), "normal", 46))
            if count == 0:
                body.add_widget(self.label("Aucun résultat.", 12, True))
        search.bind(text=rebuild)
        rebuild()

    def edit_employee(self, row, *_):
        self.employee_form(row, "Modifier employé", lambda: self.save_excel_safely(), True)

    def save_excel_safely(self):
        if not self.file_path: return
        try: self.save_excel()
        except Exception: pass

    def note(self, row, year):
        try: return float(row.get("note_" + str(year)))
        except Exception: return None

    def average(self, row, year):
        vals = [self.note(row, year - 3), self.note(row, year - 2), self.note(row, year - 1)]
        return None if any(v is None for v in vals) else round(sum(vals) / 3, 2)

    def sanction_records(self, row):
        """Return all recorded 2ème/3ème degree sanctions as (year, degree).
        The current Sanction/Année fields are also included for compatibility.
        History format accepted: 2024:2ème degré; 2026:3ème degré
        """
        records = []
        hist = str(row.get("Historique des sanctions", "") or "").strip()
        for part in hist.replace("\n", ";").split(";"):
            part = part.strip()
            if not part or ":" not in part:
                continue
            y, deg = part.split(":", 1)
            y = y.strip()
            deg = deg.strip()
            if y.isdigit() and len(y) == 4 and ("2" in norm(deg) or "3" in norm(deg)):
                records.append((int(y), deg))
        raw_y = str(row.get("Année de la sanction", "") or "").strip()
        current = str(row.get("Sanction", "") or "").strip()
        if raw_y.isdigit() and len(raw_y) == 4 and ("2" in norm(current) or "3" in norm(current)):
            rec = (int(raw_y), current)
            if rec not in records:
                records.append(rec)
        return sorted(records, key=lambda x: x[0])

    def sanction_key(self, year, degree):
        return f"{int(year)}|{norm(degree)}"

    def treated_sanctions(self, row):
        raw = str(row.get("Sanctions traitees", "") or "")
        return {x.strip() for x in raw.replace("\n", ";").split(";") if x.strip()}

    def sanction_applies(self, row, promotion_year):
        """A sanction is considered only in the 3 years preceding the target promotion.
        A sanction already used for a previous promotion calculation is not applied again.
        """
        target = int(promotion_year)
        for sy, degree in self.sanction_records(row):
            if target - 3 <= sy <= target - 1:
                if self.sanction_key(sy, degree) not in self.treated_sanctions(row):
                    return True
        return False

    def sanction_to_apply(self, row, promotion_year):
        target = int(promotion_year)
        treated = self.treated_sanctions(row)
        for sy, degree in self.sanction_records(row):
            if target - 3 <= sy <= target - 1:
                if self.sanction_key(sy, degree) not in treated:
                    return (sy, degree)
        return None

    def category_match(self, row, wanted):
        return wanted == "Toutes" or norm(row.get("Catégorie")) == norm(wanted)

    def last_year(self, row):
        return date_year(row.get("Dernière promotion"))

    def calcul(self, *_):
        if not self.rows:
            self.popup("Information", self.label("Chargez d'abord un fichier Excel ou ajoutez des employés.")); return
        box = GridLayout(cols=2, spacing=dp(6), padding=dp(8), size_hint_y=None); box.bind(minimum_height=box.setter('height'))
        year = TextInput(text=str(self.current_year), multiline=False, input_filter='int', size_hint_y=None, height=dp(42))
        minimum = TextInput(text='13', multiline=False, input_filter='float', size_hint_y=None, height=dp(42))
        cat = Spinner(text='Toutes', values=CATEGORIES, size_hint_y=None, height=dp(42))
        sanction = Spinner(text='Exclure', values=("Exclure", "Autoriser avec majoration"), size_hint_y=None, height=dp(42))
        for n, w in [("Année de promotion", year), ("Note minimale", minimum), ("Catégorie", cat), ("Sanction disciplinaire", sanction)]:
            box.add_widget(self.label(n, 12, True)); box.add_widget(w)
        rule = self.label("Condition d'ancienneté : l'année choisie - dernière promotion doit être au moins égale à 3 ans. Les candidats ne remplissant pas cette condition sont exclus.", 11)
        box.add_widget(rule); box.add_widget(Label())
        go = self.btn("Calculer", lambda *_: None, "green", 50); box.add_widget(Label()); box.add_widget(go)
        p = self.popup("Paramètres de calcul", box, (0.96, 0.82))
        def run(*_):
            try:
                y = int(year.text); mn = float(minimum.text)
                if y < 1900: raise ValueError("Année de promotion invalide.")
                results, quota, stats, non_benef = self.make_results(y, mn, cat.text, sanction.text)
                self.last_non_beneficiaries = non_benef
                self.last_results = results
                self.save_results_history(y, results)
                self.save_database()
                p.dismiss(); self.show_results(results, quota, stats, y)
            except Exception as e:
                self.popup("Erreur", self.label(str(e)))
        go.bind(on_release=run)

    def make_results(self, year, minimum, category, sanction_mode):
        eligible_by_cat = {}
        non_benef = []
        excluded_seniority = 0; excluded_note = 0; excluded_sanction = 0; excluded_missing = 0

        def add_excluded(row, avg, gap, reason, sanction_item=None):
            non_benef.append({
                "Année de promotion": year,
                "Matricule": row.get("Matricule", ""),
                "Nom et Prénom": row.get("Nom et Prénom", ""),
                "Catégorie": row.get("Catégorie", ""),
                "Dernière promotion": row.get("Dernière promotion", ""),
                "Moyenne": "" if avg is None else avg,
                "Ancienneté": "" if gap is None else gap,
                "Sanction": sanction_item[1] if sanction_item else row.get("Sanction", ""),
                "Année de la sanction": sanction_item[0] if sanction_item else row.get("Année de la sanction", ""),
                "Motif de non-bénéfice": reason
            })

        for row in self.rows:
            if not self.category_match(row, category):
                continue
            avg = self.average(row, year)
            if avg is None:
                excluded_missing += 1
                add_excluded(row, avg, None, "Notes insuffisantes ou absentes pour les 3 années précédant la promotion.")
                continue
            if avg < minimum:
                excluded_note += 1
                add_excluded(row, avg, None, f"Moyenne inférieure à la note minimale ({minimum}/20).")
                continue
            sanction_item = self.sanction_to_apply(row, year)
            bad = sanction_item is not None
            if bad and sanction_mode == "Exclure":
                excluded_sanction += 1
                row.setdefault("Sanctions traitees", "")
                treated = self.treated_sanctions(row)
                treated.add(self.sanction_key(sanction_item[0], sanction_item[1]))
                row["Sanctions traitees"] = ";".join(sorted(treated))
                add_excluded(row, avg, None, "Sanction disciplinaire 2ème/3ème degré dans les 3 années précédant la promotion.", sanction_item)
                continue
            last = self.last_year(row)
            if last is None:
                excluded_missing += 1
                add_excluded(row, avg, None, "Dernière promotion manquante.", sanction_item)
                continue
            gap = year - last
            delay = 0
            if bad and sanction_mode == "Autoriser avec majoration":
                delay = 2 if "3" in norm(sanction_item[1]) else 1
                row.setdefault("Sanctions traitees", "")
                treated = self.treated_sanctions(row)
                treated.add(self.sanction_key(sanction_item[0], sanction_item[1]))
                row["Sanctions traitees"] = ";".join(sorted(treated))
            if gap < 3 + delay:
                excluded_seniority += 1
                reason = f"Ancienneté insuffisante : {gap} an(s), minimum requis {3 + delay} an(s)."
                add_excluded(row, avg, gap, reason, sanction_item if bad else None)
                continue
            cat_key = str(row.get("Catégorie", "")).strip() or "Sans catégorie"
            eligible_by_cat.setdefault(cat_key, []).append((row, avg, gap))

        def dt_value(v, fallback):
            s = str(v or "")
            return s if s else fallback

        def rounded_quota(n, pct):
            # Règle demandée : 0,50 et plus est arrondi à l'entier supérieur.
            return int(math.floor(n * pct + 0.5))

        results = []
        quota_by_category = {}
        pace_order = {"Rapide": 0, "Moyen": 1, "Lent": 2}
        for cat_name, eligible in eligible_by_cat.items():
            # Classement à l'intérieur de chaque catégorie : notation, ancienneté, âge.
            eligible.sort(key=lambda x: (-x[1], -x[2], dt_value(x[0].get('Date de naissance'), '9999')))
            n = len(eligible)
            fast = rounded_quota(n, 0.50)
            medium = rounded_quota(n, 0.30)
            slow = rounded_quota(n, 0.20)
            # Les arrondis sont indépendants, mais on ne peut jamais attribuer plus de places que de candidats.
            if fast > n: fast = n
            remaining = n - fast
            if medium > remaining: medium = remaining
            remaining -= medium
            if slow > remaining: slow = remaining
            quota_by_category[cat_name] = {"total": n, "Rapide": fast, "Moyen": medium, "Lent": slow}
            for i, (row, avg, gap) in enumerate(eligible):
                pace = "Rapide" if i < fast else ("Moyen" if i < fast + medium else "Lent")
                next_year = year + {"Rapide": 3, "Moyen": 4, "Lent": 5}[pace]
                results.append({
                    "Matricule": row.get("Matricule", ""), "Nom et Prénom": row.get("Nom et Prénom", ""),
                    "Catégorie": row.get("Catégorie", ""), "Dernière promotion": row.get("Dernière promotion", ""),
                    "Moyenne": avg, "Ancienneté": gap, "Rythme proposé": pace,
                    "Prochaine promotion": next_year, "Sanction": row.get("Sanction", "")
                })

        # Affichage final : catégorie, rythme, puis notation décroissante.
        results.sort(key=lambda r: (norm(r.get("Catégorie")), pace_order.get(r.get("Rythme proposé"), 9), -float(r.get("Moyenne") or 0), -int(r.get("Ancienneté") or 0), norm(r.get("Nom et Prénom"))))
        total = len(results)
        quota = {"total": total, "categories": quota_by_category}
        stats = {"excl_seniority": excluded_seniority, "excl_note": excluded_note, "excl_sanction": excluded_sanction, "excl_missing": excluded_missing}
        return results, quota, stats, non_benef

    def show_results(self, results, quota, stats, year):
        box = BoxLayout(orientation='vertical', spacing=dp(5), padding=dp(6))
        box.add_widget(self.label(f"Année de promotion : {year}", 14, True))
        qparts = []
        for cat_name, q in quota.get("categories", {}).items():
            qparts.append(f"{cat_name}: {q['Rapide']} Rapide / {q['Moyen']} Moyen / {q['Lent']} Lent")
        box.add_widget(self.label("Quotas par catégorie : " + (" | ".join(qparts) if qparts else "Aucun candidat éligible"), 10, True))
        box.add_widget(self.label(f"Éligibles : {quota['total']} | Exclus : ancienneté {stats['excl_seniority']} | note {stats['excl_note']} | sanction {stats['excl_sanction']} | données manquantes {stats['excl_missing']}", 10))
        if not results:
            box.add_widget(self.label("Aucun candidat éligible pour cette année.", 13, True))
        else:
            scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
            grid = GridLayout(cols=9, size_hint_y=None, row_default_height=dp(42), spacing=dp(1)); grid.bind(minimum_height=grid.setter('height'))
            heads = ['#','Matricule','Nom et Prénom','Catégorie','Dernière promotion','Moyenne','Ancienneté','Rythme','Prochaine promotion']
            for h in heads: grid.add_widget(self.label(h, 9, True))
            for i, r in enumerate(results, 1):
                vals = [i, r['Matricule'], r['Nom et Prénom'], r['Catégorie'], r['Dernière promotion'], r['Moyenne'], str(r['Ancienneté']) + ' ans', r['Rythme proposé'], r['Prochaine promotion']]
                for v in vals: grid.add_widget(self.label(v, 9))
            scroll.add_widget(grid); box.add_widget(scroll)
        bar = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(5))
        ex = self.btn('Exporter Excel', lambda *_: self.export_results(results, year), 'orange')
        cl = self.btn('Fermer', lambda *_: None, 'gray'); bar.add_widget(ex); bar.add_widget(cl); box.add_widget(bar)
        p = self.popup('Résultats de promotion', box, (0.99, 0.92)); cl.bind(on_release=lambda *_: p.dismiss())

    def read_results_history(self):
        if not os.path.exists(RESULTS_HISTORY_FILE):
            return []
        try:
            with open(RESULTS_HISTORY_FILE, newline="", encoding="utf-8-sig") as f:
                return list(csv.DictReader(f))
        except Exception:
            return []

    def write_results_history(self, data):
        os.makedirs(APP_DIR, exist_ok=True)
        with open(RESULTS_HISTORY_FILE, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=RESULT_HISTORY_FIELDS)
            w.writeheader()
            w.writerows(data)

    def save_results_history(self, year, results):
        """Automatically archive the latest calculation for each promotion year.
        Recalculating the same year replaces that year's previous result instead of duplicating it.
        """
        try:
            data = self.read_results_history()
            y = str(year)
            data = [r for r in data if str(r.get("Année de calcul", "")) != y]
            for r in results:
                item = {k: "" for k in RESULT_HISTORY_FIELDS}
                item["Année de calcul"] = y
                for k in RESULT_HISTORY_FIELDS:
                    if k == "Année de calcul":
                        continue
                    item[k] = r.get(k, "")
                data.append(item)
            data.sort(key=lambda r: (int(r.get("Année de calcul", "0") or 0), norm(r.get("Nom et Prénom", ""))))
            self.write_results_history(data)
            return True
        except Exception:
            return False

    def saved_results_history(self, *_):
        data = self.read_results_history()
        outer = BoxLayout(orientation="vertical", spacing=dp(6), padding=dp(6))
        search = TextInput(hint_text="Rechercher par année, nom ou matricule...", multiline=False, size_hint_y=None, height=dp(42))
        outer.add_widget(search)
        info = self.label("Les résultats sont enregistrés automatiquement après chaque calcul.", 10)
        outer.add_widget(info)
        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        grid = GridLayout(cols=len(RESULT_HISTORY_FIELDS), size_hint_y=None, row_default_height=dp(42), spacing=dp(1), padding=dp(2))
        grid.bind(minimum_height=grid.setter("height"))
        scroll.add_widget(grid); outer.add_widget(scroll)
        close = self.btn("Fermer", lambda *_: None, "gray", 48); outer.add_widget(close)
        pop = self.popup("Historique des résultats de promotion", outer, (0.99, 0.94))
        close.bind(on_release=lambda *_: pop.dismiss())

        def rebuild(*_):
            grid.clear_widgets()
            for h in RESULT_HISTORY_FIELDS:
                grid.add_widget(self.label(h, 8, True))
            q = norm(search.text)
            shown = 0
            for item in reversed(data):
                joined = norm(" ".join(str(item.get(k, "")) for k in ("Année de calcul", "Matricule", "Nom et Prénom")))
                if q and q not in joined:
                    continue
                shown += 1
                for k in RESULT_HISTORY_FIELDS:
                    grid.add_widget(self.label(item.get(k, ""), 8))
            info.text = f"Lignes enregistrées : {shown} | Une nouvelle simulation remplace uniquement l'ancienne simulation de la même année."
            if not shown:
                for _ in RESULT_HISTORY_FIELDS:
                    grid.add_widget(self.label("Aucun résultat enregistré.", 9, True))
        search.bind(text=rebuild)
        rebuild()

    def export_results(self, results, year=None):
        try:
            if load_workbook is None:
                raise Exception("openpyxl est nécessaire pour exporter en Excel.")
            os.makedirs(APP_DIR, exist_ok=True)
            wb = __import__('openpyxl').Workbook()
            ws = wb.active
            ws.title = "Résultats promotion"
            fields = ['Matricule','Nom et Prénom','Catégorie','Dernière promotion','Moyenne','Ancienneté','Rythme proposé','Prochaine promotion','Sanction']
            ws.append(fields)
            for r in results:
                ws.append([r.get(k, '') for k in fields])

            ws2 = wb.create_sheet("Non bénéficiaires")
            ws2.append(NON_BENEF_FIELDS)
            for item in getattr(self, 'last_non_beneficiaries', []):
                ws2.append([item.get(k, '') for k in NON_BENEF_FIELDS])

            # Mise en forme simple et largeur lisible.
            for sheet in (ws, ws2):
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = sheet.dimensions
                for col in sheet.columns:
                    letter = col[0].column_letter
                    max_len = max(len(str(c.value or '')) for c in col)
                    sheet.column_dimensions[letter].width = min(max(max_len + 2, 10), 35)

            wb.save(RESULT_FILE)
            self.save_database()
            self.popup('Export', self.label('Fichier Excel enregistré dans :\n' + RESULT_FILE + '\n\nFeuilles : Résultats promotion et Non bénéficiaires.'))
        except Exception as e:
            self.popup('Erreur', self.label(str(e)))

    def export_last(self, *_):
        self.export_results(self.last_results, self.current_year)

    def notations(self, *_):
        if not self.rows:
            self.popup('Information', self.label("Chargez d'abord un fichier Excel.")); return
        outer = BoxLayout(orientation='vertical', spacing=dp(5), padding=dp(5))
        top = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(5))
        search = TextInput(hint_text="Rechercher par matricule ou nom...", multiline=False)
        top.add_widget(search)
        year_input = TextInput(hint_text="Année", text=str(self.current_year), multiline=False, input_filter='int', size_hint_x=None, width=dp(90))
        top.add_widget(year_input)
        add_year = self.btn("+ Année", lambda *_: None, "green", 44); top.add_widget(add_year)
        outer.add_widget(top)
        hint = self.label("Ajoutez une année manuellement. Elle sera disponible pour tous les employés.", 10)
        outer.add_widget(hint)
        scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        body = GridLayout(cols=6, size_hint_y=None, row_default_height=dp(44), spacing=dp(2), padding=dp(3)); body.bind(minimum_height=body.setter('height'))
        scroll.add_widget(body); outer.add_widget(scroll)
        close = self.btn('Fermer', lambda *_: None, 'gray', 48); outer.add_widget(close)
        p = self.popup('Gérer les notations', outer, (0.99, 0.94)); close.bind(on_release=lambda *_: p.dismiss())

        def note_years():
            return all_note_years(self.rows)

        def rebuild(*_):
            years = note_years()
            # Display two note columns plus an action column, keeping the mobile table readable.
            body.clear_widgets()
            headers = ["Matricule", "Nom et Prénom", "Catégorie", "Notes", "Années", "Action"]
            for h in headers: body.add_widget(self.label(h, 9, True))
            q = norm(search.text)
            for r in self.rows:
                txt = norm(str(r.get('Matricule','')) + ' ' + str(r.get('Nom et Prénom','')))
                if q and q not in txt: continue
                vals = []
                for y in years:
                    v = r.get('note_'+str(y), '')
                    if v not in ('', None): vals.append(f"{y}: {v}")
                notes_txt = " | ".join(vals) if vals else "—"
                body.add_widget(self.label(r.get('Matricule',''), 9))
                body.add_widget(self.label(r.get('Nom et Prénom',''), 9))
                body.add_widget(self.label(r.get('Catégorie',''), 9))
                body.add_widget(self.label(notes_txt, 9))
                body.add_widget(self.label(str(len(years)) + " année(s)", 9))
                body.add_widget(self.btn('Modifier', lambda inst, row=r: self.edit_notes(row), 'green', 38))

        def add_year_action(*_):
            try:
                y = int(year_input.text)
                if y < 1900 or y > 2200: raise ValueError
                key = 'note_' + str(y)
                for r in self.rows:
                    if key not in r: r[key] = ''
                self.current_year = y
                self.save_database(); self.save_excel_safely(); rebuild()
                self.popup('Année ajoutée', self.label(f"L'année {y} a été ajoutée aux notations."))
            except Exception:
                self.popup('Erreur', self.label("Veuillez saisir une année valide."))
        add_year.bind(on_release=add_year_action)
        search.bind(text=rebuild); rebuild()

    def edit_notes(self, row):
        years = all_note_years(self.rows)
        outer = BoxLayout(orientation='vertical')
        scroll = ScrollView(); form = GridLayout(cols=2, spacing=dp(5), padding=dp(8), size_hint_y=None); form.bind(minimum_height=form.setter('height'))
        fields = {}
        for y in years:
            form.add_widget(self.label('Note ' + str(y), 11, True))
            f = TextInput(text='' if row.get('note_'+str(y)) is None else str(row.get('note_'+str(y))), multiline=False, input_filter='float', size_hint_y=None, height=dp(42))
            fields[y] = f; form.add_widget(f)
        scroll.add_widget(form); outer.add_widget(scroll)
        bar=BoxLayout(size_hint_y=None,height=dp(50),spacing=dp(5)); save=self.btn('Enregistrer',lambda *_:None,'green'); close=self.btn('Fermer',lambda *_:None,'gray'); bar.add_widget(save);bar.add_widget(close);outer.add_widget(bar)
        p=self.popup('Notation : '+str(row.get('Nom et Prénom','')),outer,(0.95,0.90)); close.bind(on_release=lambda *_:p.dismiss())
        def sv(*_):
            for y,f in fields.items(): row['note_'+str(y)] = f.text
            self.save_database(); self.save_excel_safely(); p.dismiss()
        save.bind(on_release=sv)

    def save_excel(self):
        if not self.file_path: raise Exception('Aucun fichier Excel chargé.')
        wb = load_workbook(self.file_path); ws = wb.active
        cols = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
        # Add missing notation columns created manually in the application.
        years = all_note_years(self.rows)
        for y in years:
            key = 'note_' + str(y)
            if key not in cols:
                new_col = ws.max_column + 1
                ws.cell(1, new_col).value = key
                cols[key] = new_col
        for idx, row in enumerate(self.rows, 2):
            normalize_row_dates(row)
            for key, col in cols.items():
                if key in row: ws.cell(idx, col).value = row.get(key, '')
        wb.save(self.file_path)

    def read_history(self):
        if not os.path.exists(HISTORY_FILE): return []
        with open(HISTORY_FILE, newline='', encoding='utf-8-sig') as f: return list(csv.DictReader(f))

    def write_history(self, data):
        os.makedirs(APP_DIR, exist_ok=True)
        with open(HISTORY_FILE, 'w', newline='', encoding='utf-8-sig') as f:
            w=csv.DictWriter(f, fieldnames=HISTORY_FIELDS); w.writeheader(); w.writerows(data)

    def history(self, *_):
        # Liste des travailleurs, présentée comme « Modifier sélection ».
        outer = BoxLayout(orientation='vertical', spacing=dp(6), padding=dp(6))
        search = TextInput(hint_text="Rechercher par nom ou matricule...", multiline=False,
                           size_hint_y=None, height=dp(42))
        outer.add_widget(search)
        info = self.label("Sélectionnez un travailleur pour afficher toutes ses promotions.", 11)
        outer.add_widget(info)

        scroll = ScrollView(do_scroll_x=False, do_scroll_y=True)
        body = GridLayout(cols=1, spacing=dp(4), padding=dp(4), size_hint_y=None)
        body.bind(minimum_height=body.setter('height'))
        scroll.add_widget(body); outer.add_widget(scroll)
        close = self.btn('Fermer', lambda *_: None, 'gray', 48)
        outer.add_widget(close)
        p = self.popup('Gérer historique promotions', outer, (0.97, 0.94))
        close.bind(on_release=lambda *_: p.dismiss())

        def employee_history_popup(row):
            # Affiche toutes les lignes d'historique liées au travailleur.
            outer2 = BoxLayout(orientation='vertical', spacing=dp(6), padding=dp(6))
            head = Card(orientation='vertical', size_hint_y=None, height=dp(82))
            head.add_widget(self.label(str(row.get('Nom et Prénom', '')), 15, True))
            head.add_widget(self.label("Matricule : " + str(row.get('Matricule', '')), 11))
            outer2.add_widget(head)

            sc = ScrollView(do_scroll_x=True, do_scroll_y=True)
            b = GridLayout(cols=7, size_hint_y=None, row_default_height=dp(43), spacing=dp(2), padding=dp(3))
            b.bind(minimum_height=b.setter('height'))
            sc.add_widget(b); outer2.add_widget(sc)

            actions = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(5))
            new = self.btn('Ajouter promotion', lambda *_: None, 'green')
            done = self.btn('Fermer', lambda *_: None, 'gray')
            actions.add_widget(new); actions.add_widget(done); outer2.add_widget(actions)
            hp = self.popup('Historique de ' + str(row.get('Nom et Prénom', '')), outer2, (0.99, 0.92))
            done.bind(on_release=lambda *_: hp.dismiss())

            def rebuild_h(*_):
                b.clear_widgets()
                for h in HISTORY_FIELDS:
                    b.add_widget(self.label(h, 9, True))
                # Colonne d'action séparée.
                b.add_widget(self.label('Action', 9, True))
                data = self.read_history()
                mat = norm(row.get('Matricule', ''))
                name = norm(row.get('Nom et Prénom', ''))
                found = []
                for i, item in enumerate(data):
                    imat = norm(item.get('Matricule', ''))
                    iname = norm(item.get('Nom et Prénom', ''))
                    if (mat and imat == mat) or (not mat and name and iname == name):
                        found.append((i, item))
                # Trier par année croissante pour voir toutes les promotions dans l'ordre.
                def yr(x):
                    try: return int(str(x[1].get('Année', '')).strip())
                    except Exception: return 9999
                found.sort(key=yr)
                for idx, item in found:
                    for key in HISTORY_FIELDS:
                        b.add_widget(self.label(item.get(key, ''), 9))
                    b.add_widget(self.btn('Modifier', lambda inst, i=idx: edit_history(i), 'orange', 38))
                if not found:
                    b.add_widget(self.label('Aucune promotion enregistrée pour ce travailleur.', 11, True))
                    for _ in range(7): b.add_widget(self.label('', 9))

            def edit_history(index=None):
                data = self.read_history()
                if index is not None and 0 <= index < len(data):
                    item = data[index].copy()
                else:
                    item = {k: '' for k in HISTORY_FIELDS}
                    item['Matricule'] = row.get('Matricule', '')
                    item['Nom et Prénom'] = row.get('Nom et Prénom', '')
                    item['Rythme'] = 'Rapide'

                formbox = BoxLayout(orientation='vertical', spacing=dp(5))
                scform = ScrollView(do_scroll_x=False)
                form = GridLayout(cols=2, spacing=dp(5), padding=dp(8), size_hint_y=None)
                form.bind(minimum_height=form.setter('height'))
                fields = {}
                for key in HISTORY_FIELDS:
                    form.add_widget(self.label(key, 11, True))
                    if key == 'Rythme':
                        w = Spinner(text=item.get(key, '') or 'Rapide', values=('Rapide', 'Moyen', 'Lent'),
                                     size_hint_y=None, height=dp(42))
                    else:
                        w = TextInput(text='' if item.get(key) is None else str(item.get(key)),
                                      multiline=False, size_hint_y=None, height=dp(42))
                    fields[key] = w; form.add_widget(w)
                scform.add_widget(form); formbox.add_widget(scform)

                bar = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(5))
                sv = self.btn('Enregistrer', lambda *_: None, 'green')
                cn = self.btn('Annuler', lambda *_: None, 'gray')
                bar.add_widget(sv); bar.add_widget(cn); formbox.add_widget(bar)
                ep = self.popup('Modifier promotion' if index is not None else 'Ajouter promotion',
                                formbox, (0.96, 0.90))
                cn.bind(on_release=lambda *_: ep.dismiss())

                def save_h(*_):
                    val = {}
                    for k in HISTORY_FIELDS:
                        w = fields[k]
                        val[k] = w.text if hasattr(w, 'text') else str(w.text)
                    val['Matricule'] = val.get('Matricule') or str(row.get('Matricule', ''))
                    val['Nom et Prénom'] = val.get('Nom et Prénom') or str(row.get('Nom et Prénom', ''))
                    val["Date d'effet"] = format_date(val.get("Date d'effet"))
                    if index is None:
                        data.append(val)
                    else:
                        data[index] = val
                    self.write_history(data)
                    self.save_database()
                    ep.dismiss()
                    rebuild_h()

                sv.bind(on_release=save_h)

            new.bind(on_release=lambda *_: edit_history(None))
            rebuild_h()

        def rebuild(*_):
            body.clear_widgets()
            q = norm(search.text)
            count = 0
            for row in self.rows:
                joined = norm(str(row.get('Matricule', '')) + ' ' + str(row.get('Nom et Prénom', '')))
                if q and q not in joined:
                    continue
                # Présentation identique à Modifier sélection : une fiche cliquable par travailleur.
                card = Card(orientation='horizontal', size_hint_y=None, height=dp(68), spacing=dp(6))
                left = BoxLayout(orientation='vertical')
                left.add_widget(self.label(str(row.get('Nom et Prénom', '')), 13, True))
                left.add_widget(self.label('Matricule : ' + str(row.get('Matricule', '')), 10))
                left.add_widget(self.label(str(row.get('Catégorie', '')), 10))
                card.add_widget(left)
                open_btn = self.btn('Ouvrir', lambda inst, r=row: employee_history_popup(r), 'orange', 48)
                card.add_widget(open_btn)
                body.add_widget(card)
                count += 1
            info.text = f"Travailleurs affichés : {count} | Cliquez sur Ouvrir pour voir toutes les promotions."
            if count == 0:
                body.add_widget(self.label('Aucun employé.', 12, True))

        search.bind(text=rebuild)
        rebuild()

    def save_database(self):
        """Save the complete working database as JSON, with date values free of time."""
        try:
            os.makedirs(APP_DIR, exist_ok=True)
            rows = []
            for src in self.rows:
                row = dict(src)
                normalize_row_dates(row)
                rows.append(row)
            history = []
            for src in self.read_history():
                item = dict(src)
                item["Date d'effet"] = format_date(item.get("Date d'effet"))
                history.append(item)
            result_history = [dict(x) for x in self.read_results_history()]
            payload = {
                "version": 2,
                "application": "Avancement Horizontal",
                "saved_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "current_year": self.current_year,
                "rows": rows,
                "history": history,
                "result_history": result_history
            }
            with open(DATABASE_FILE, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def load_database_safely(self):
        """Load the automatic local JSON database if it exists."""
        if not os.path.exists(DATABASE_FILE):
            return
        try:
            with open(DATABASE_FILE, encoding="utf-8") as f:
                data = json.load(f)
            self.rows = data.get("rows", []) or []
            for row in self.rows:
                normalize_row_dates(row)
            history = data.get("history", []) or []
            self.write_history(history)
            result_history = data.get("result_history", []) or []
            self.write_results_history(result_history)
            if data.get("current_year"):
                self.current_year = int(data.get("current_year"))
            self.last_results = []
        except Exception:
            self.rows = []
            self.last_results = []

    def save_backup(self, *_):
        """Create a dated JSON backup that can be copied/downloaded later."""
        if not self.rows and not self.read_history():
            self.popup("Information", self.label("Aucune donnée à sauvegarder."))
            return
        self.save_database()
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(APP_DIR, BACKUP_PREFIX + stamp + ".json")
        try:
            shutil.copy2(DATABASE_FILE, path)
            self.popup("Sauvegarder", self.label("Copie JSON créée avec succès :\n\n" + path + "\n\nVous pouvez conserver ce fichier et le restaurer plus tard."))
        except Exception as e:
            self.popup("Erreur", self.label(str(e)))

    def restore_backup(self, *_):
        box = BoxLayout(orientation="vertical")
        chooser = __import__('kivy.uix.filechooser', fromlist=['FileChooserListView']).FileChooserListView(path=APP_DIR if os.path.isdir(APP_DIR) else "/storage/emulated/0", filters=["*.json"])
        box.add_widget(chooser)
        bar = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(5))
        load = self.btn("Restaurer", lambda *_: None, "green")
        cancel = self.btn("Annuler", lambda *_: None, "gray")
        bar.add_widget(load); bar.add_widget(cancel); box.add_widget(bar)
        p = self.popup("Restaurer une sauvegarde JSON", box, (0.98, 0.88))
        def do_restore(*_):
            if not chooser.selection:
                self.popup("Information", self.label("Sélectionnez un fichier JSON."))
                return
            try:
                with open(chooser.selection[0], encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, dict) or "rows" not in data:
                    raise ValueError("Fichier JSON de sauvegarde invalide.")
                self.rows = data.get("rows", []) or []
                for row in self.rows:
                    normalize_row_dates(row)
                history = data.get("history", []) or []
                self.write_history(history)
                result_history = data.get("result_history", []) or []
                self.write_results_history(result_history)
                self.last_results = []
                if data.get("current_year"):
                    self.current_year = int(data.get("current_year"))
                self.save_database()
                p.dismiss(); self.show_home()
                self.popup("Restaurer", self.label("Sauvegarde restaurée avec succès."))
            except Exception as e:
                self.popup("Erreur", self.label(str(e)))
        load.bind(on_release=do_restore); cancel.bind(on_release=lambda *_: p.dismiss())

    def clear_list(self, *_):
        self.rows=[];self.last_results=[];self.file_path=''
        self.save_database()
        self.show_home()


if __name__=='__main__': MainApp().run()
