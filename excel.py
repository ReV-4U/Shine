# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Génération des classeurs Excel (XlsxWriter).

Structure d'un classeur :
  En-tête | Instructions | Synthèse | Récapitulatif | <une feuille par ressource> | _calc, _meta (masqués)

Une ligne de feuille = une ligne d'un CSV (aucune fusion Full/Read, aucun dédoublonnage).

Contrat de relecture (phase 2, commande `ingest`) :
  * jamais de cellule fusionnée dans les zones de données (seul le commentaire général, sous le
    tableau de chaque feuille de ressource, est fusionné) ;
  * chaque ligne porte sa clé stable dans une colonne masquée « Clé » : le tri et le filtre de
    l'utilisateur ne cassent donc pas la correspondance ;
  * l'onglet _meta décrit, par feuille, où se trouvent le tableau, le bloc d'ajout et le
    commentaire général (voir README).
"""
from __future__ import annotations

import datetime as dt
import hashlib
import math
import re
from dataclasses import dataclass, field

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from model import FULL, READ, Book, ResourceData, ScopeData

SCHEMA_VERSION = 4
TOOL_VERSION = "2.2.0"

LEVEL_LABELS = {FULL: "Modification", READ: "Lecture"}
ADD_COLS = ("Nom", "Identifiant (optionnel)", "Groupe de ressource", "Justification")   # tableau « Personnes à ajouter » (Synthèse)
ADD_GROUP_SPAN = (2, 4)      # colonnes C:E fusionnées par ligne (liste « Groupe de ressource »), valeur lue en C
ADD_JUSTIF_SPAN = (5, 7)     # colonnes F:H fusionnées par ligne, valeur lue en F
# colonnes de l'onglet _meta (contrat phase 2) ; les blocs sont placés côte à côte
META_SHEETS = ["sheet", "folder", "subpath", "resource", "group_full", "group_read", "header_row", "first_data_row",
               "last_data_row", "n_rows", "comment_cell", "decision_cols", "key_col", "data_hash"]
META_ROWS = ["sheet", "key", "sam", "level", "initial_row"]
META_GROUPS = ["group", "sheet", "folder", "subpath", "resource", "level", "file", "owner"]
META_C0 = 3
META_C1 = META_C0 + len(META_SHEETS) + 1
META_C2 = META_C1 + len(META_ROWS) + 1          # colonne de la liste « group » (source de la liste déroulante)
SYN_HEAD_ROW = 3                                 # en-tête (0-based) du tableau de la Synthèse


def syn_layout(n_resources: int) -> tuple[int, int, int, int]:
    """(première ligne de données, dernière, ligne de total, titre du tableau « Personnes à ajouter »), 0-based."""
    first = SYN_HEAD_ROW + 1
    last = SYN_HEAD_ROW + max(n_resources, 1)
    return first, last, last + 1, last + 4
DROPDOWN_VALUES = ["Oui", "Non"]
SHEET_COVER, SHEET_INSTR, SHEET_SYN, SHEET_RECAP, SHEET_CALC = "En-tête", "Instructions", "Synthèse", "Récapitulatif", "_calc"
RECAP_HEAD_ROW = 6          # en-tête (0-based) du tableau de l'onglet Récapitulatif
RECAP_COLS = ("Action", "Propriétaire du groupe", "Nom", "Identifiant", "Droit souhaité", "Justification", "Ressource")
REMOVAL_RIGHT = "Aucun (retrait)"
CLASSIFICATIONS = ("1-Public", "2-Restreint", "3-Confidentiel")          # niveaux de l'En-tête (review.classification)
UPDATE_COLS = ("Auteur", "Date", "Version", "Nature de la modification", "Approbation", "Validation", "Status")
NO_MEMBER = "Aucun membre"

_BAD_SHEET = re.compile(r"[\[\]:*?/\\]")
_RESERVED = {"instructions", "synthèse", "synthese", "_meta", "_calc", "history", "en-tête", "en-tete", "entête",
             "récapitulatif", "recapitulatif"}
_HEADER_ROW = 5          # ligne (0-based) de l'en-tête du tableau dans une feuille de ressource


def safe(value) -> str:
    """Neutralise l'injection de formules : préfixe ' si la valeur commence par = + - @."""
    s = "" if value is None else str(value)
    return "'" + s if s and s[0] in "=+-@\t\r" else s


def _q(s: str) -> str:
    return s.replace('"', '""')


def _sq(s: str) -> str:
    return s.replace("'", "''")


def make_sheet_name(base: str, used: set[str]) -> str:
    """Nom d'onglet valide : ≤ 31 caractères, sans caractères interdits, unique (casse ignorée)."""
    n = _BAD_SHEET.sub("_", base).strip().strip("'").strip() or "Ressource"
    cand = n[:31].rstrip().rstrip("'") or "Ressource"
    k = 2
    while cand.casefold() in used or cand.casefold() in _RESERVED:
        tag = f"~{k}"
        cand = n[: 31 - len(tag)].rstrip().rstrip("'") + tag
        k += 1
    used.add(cand.casefold())
    return cand


@dataclass
class SheetInfo:
    name: str
    scope: ScopeData
    res: ResourceData
    first: int = 0            # première ligne de données (0-based)
    last: int = 0
    n_rows: int = 0           # nombre de lignes de données (0 si ressource vide)
    header_row: int = _HEADER_ROW
    comment_cell: str = ""
    dec_cols: tuple[int, ...] = ()
    key_col: int = 0
    data_hash: str = ""


@dataclass
class BookReport:
    path: str
    sheets: list[SheetInfo] = field(default_factory=list)
    rows: int = 0
    data_hash: str = ""


class _Formats:
    def __init__(self, wb, cfg):
        self.wb, self.cache = wb, {}
        self.base = {"font_name": cfg["excel"]["font"], "font_size": cfg["excel"]["font_size"], "valign": "vcenter"}

    def __call__(self, **kw):
        props = {**self.base, **kw}
        key = tuple(sorted(props.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(props)
        return self.cache[key]

    def dxf(self, **kw):  # formats de mise en forme conditionnelle (sans police)
        key = ("dxf",) + tuple(sorted(kw.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(kw)
        return self.cache[key]


def _group_name(g) -> str:
    return g.rel_path.rsplit("/", 1)[-1] if g else ""


def _rows_for_hash(scope: ScopeData, res: ResourceData) -> list[str]:
    out = [f"R|{scope.name}|{res.subpath}|{res.name}|{_group_name(res.full)}|{_group_name(res.read)}"]
    for p in res.rows:
        out.append("|".join((p.key, p.sam, p.name, p.dn, p.level)))
    return out


def _digest(lines: list[str]) -> str:
    return hashlib.sha256("\n".join(lines).encode("utf-8")).hexdigest()


class BookBuilder:
    def __init__(self, path, book: Book, cfg: dict, mode: str, period: str, now: dt.datetime):
        self.path, self.book, self.cfg, self.mode, self.period, self.now = path, book, cfg, mode, period, now
        self.native = mode == "native"
        self.show_dn = bool(cfg["excel"].get("show_dn"))
        self.c = cfg["excel"]["colors"]
        self.wb = xlsxwriter.Workbook(str(path), {
            "strings_to_formulas": False, "strings_to_urls": False, "strings_to_numbers": False,
            "default_date_format": "dd/mm/yyyy"})
        self.F = _Formats(self.wb, cfg)
        self.report = BookReport(str(path))
        self.used_names = set(_RESERVED)
        self.used_tables: set[str] = set()
        self.add_title = self.add_header_row = self.add_first = self.add_last = 0
        self.groups: list[dict] = []
        self.protect_opts = {"autofilter": True, "sort": True, "format_columns": True, "format_rows": True}

    # ------------------------------------------------------------------ utilitaires
    def _protect(self, ws):
        if self.cfg["excel"]["protect_sheets"]:
            ws.protect("", self.protect_opts)   # SANS mot de passe

    def _date_dv(self, title: str) -> dict:
        """Liste déroulante de dates (feuille cachée _calc, colonne Q), sans erreur bloquante à la saisie libre."""
        n = self.cfg["excel"]["date_choices_days"] + 1
        return {"validate": "list", "source": f"='{SHEET_CALC}'!$Q$2:$Q${n + 1}", "ignore_blank": True,
                "show_error": False, "input_title": title,
                "input_message": "Choisissez une date dans la liste, à partir d'aujourd'hui (ou saisissez-la)."}

    def _table_name(self, prefix: str) -> str:
        n = 1
        while f"{prefix}_{n}" in self.used_tables:
            n += 1
        self.used_tables.add(f"{prefix}_{n}")
        return f"{prefix}_{n}"

    def _hdr(self, input_col=False):
        c = self.c
        return self.F(bold=True, font_color=c["header_font"], bg_color=c["input_header_bg"] if input_col else c["header_bg"],
                      border=1, border_color=c["border"], text_wrap=True, align="center")

    # ------------------------------------------------------------------ construction
    def build(self) -> BookReport:
        wb = self.wb
        wb.set_properties({"title": f"Revue des droits – {self.book.label} – {self.period}",
                           "subject": "Revue périodique des droits d'accès", "created": self.now,
                           "comments": f"schema {SCHEMA_VERSION}"})
        ws_cover = wb.add_worksheet(SHEET_COVER)
        ws_instr = wb.add_worksheet(SHEET_INSTR)
        ws_syn = wb.add_worksheet(SHEET_SYN)
        ws_recap = wb.add_worksheet(SHEET_RECAP)

        # un nom de ressource présent dans plusieurs dossiers du classeur est suffixé par le code du dossier
        counts: dict[str, int] = {}
        for s in self.book.scopes:
            for r in s.resources:
                counts[r.name.casefold()] = counts.get(r.name.casefold(), 0) + 1
        infos: list[SheetInfo] = []
        for s in self.book.scopes:
            for r in s.resources:
                base = r.name
                if counts[r.name.casefold()] > 1:
                    tag = "_" + (s.code or s.label)
                    base = r.name[: max(1, 31 - len(tag))] + tag
                infos.append(SheetInfo(make_sheet_name(base, self.used_names), s, r))
        self.add_title = syn_layout(len(infos))[3]
        self.groups = self._groups(infos)
        for info in infos:
            self._resource_sheet(info)

        # capacité du récapitulatif : toutes les lignes pourraient être à retirer, dans la limite de la configuration
        self.n_rem = sum(i.n_rows for i in infos)
        self.cap_rem = min(self.n_rem, self.cfg["excel"]["recap_max_removals"])
        self.n_add = self.cfg["excel"]["add_rows"]

        self._synthese(ws_syn, infos)          # fixe la position du tableau « Personnes à ajouter »
        self._recap(ws_recap)
        self._instructions(ws_instr)           # fixe la position du bloc de signature (lue par l'En-tête)
        self._cover(ws_cover, infos)
        ws_calc = wb.add_worksheet(SHEET_CALC)
        self._calc_sheet(ws_calc, infos)
        ws_meta = wb.add_worksheet("_meta")
        self._meta_sheet(ws_meta, infos)
        ws_calc.hide()
        ws_meta.hide()
        ws_cover.activate()
        wb.close()

        self.report.sheets = infos
        self.report.rows = sum(i.n_rows for i in infos)
        self.report.data_hash = self._data_hash()
        return self.report

    def _groups(self, infos: list[SheetInfo]) -> list[dict]:
        """Groupes (fichiers CSV) du classeur : libellé affiché dans la liste « Groupe de ressource » + infos pour _meta.

        Libellé = nom du fichier sans extension ; préfixé du dossier (et du sous-dossier) quand le classeur en regroupe
        plusieurs, pour que deux ressources de même nom restent distinguables.
        """
        multi = len(self.book.scopes) > 1
        out, seen = [], set()
        for info in infos:
            for lvl, g in ((FULL, info.res.full), (READ, info.res.read)):
                if g is None:
                    continue
                parts = ([info.scope.name] if multi else []) + ([info.res.subpath] if info.res.subpath else []) + [g.stem]
                label = base = " / ".join(parts)
                k = 2
                while label.casefold() in seen:
                    label, k = f"{base} ({k})", k + 1
                seen.add(label.casefold())
                out.append({"group": label, "sheet": info.name, "folder": info.scope.name, "subpath": info.res.subpath,
                            "resource": info.res.name, "level": lvl, "file": g.rel_path.rsplit("/", 1)[-1],
                            "owner": info.scope.label})
        return out

    def _data_hash(self) -> str:
        lines: list[str] = []
        for s in sorted(self.book.scopes, key=lambda x: x.name.casefold()):
            for r in s.resources:
                lines += _rows_for_hash(s, r)
        return _digest(lines)

    # ------------------------------------------------------------------ feuille de ressource
    def _resource_sheet(self, info: SheetInfo):
        wb, F, c, native = self.wb, self.F, self.c, self.native
        res, scope = info.res, info.scope
        ws = wb.add_worksheet(info.name)
        ws.set_tab_color(c["header_bg"])
        ws.hide_gridlines(2)

        # ---- colonnes
        cols = ["Nom", "Identifiant"]
        c_dn = None
        if self.show_dn:
            c_dn = len(cols)
            cols.append("DN (annuaire)")
        c_right = len(cols)
        cols.append("Droit actuel")
        c_dec = len(cols)
        cols += ["Conserver", "Retirer"] if native else ["Accès autorisé ?"]
        dec_cols = (c_dec, c_dec + 1) if native else (c_dec,)
        c_com = len(cols)
        cols += ["Commentaire", "Clé"]
        c_key = c_com + 1
        info.dec_cols, info.key_col = dec_cols, c_key

        members = res.rows
        n = len(members)
        first = _HEADER_ROW + 1
        last = first + max(n, 1) - 1
        info.first, info.last, info.n_rows = first, last, n

        def width(col, lo, hi, vals):
            m = max([len(v) for v in vals] + [len(cols[col])])
            return max(lo, min(hi, round(m * 1.08 + 2)))
        ws.set_column(0, 0, width(0, 22, 46, [p.name for p in members]))
        ws.set_column(1, 1, width(1, 16, 30, [p.sam for p in members]))
        if c_dn is not None:
            ws.set_column(c_dn, c_dn, width(c_dn, 30, 70, [p.dn for p in members]))
        ws.set_column(c_right, c_right, 15)
        ws.set_column(c_dec, dec_cols[-1], 17 if not native else 12)
        ws.set_column(c_com, c_com, 42)
        ws.set_column(c_key, c_key, 18, None, {"hidden": True})

        # ---- en-tête de feuille
        lab = F(bold=True, font_color=c["title"])
        ws.write_string(0, 0, "Ressource", lab)
        ws.write_string(0, 1, safe(res.name), F(bold=True, font_size=self.cfg["excel"]["font_size"] + 5, font_color=c["title"]))
        ws.set_row(0, 24)
        ws.write_string(1, 0, "Dossier d'origine", lab)
        ws.write_string(1, 1, safe(scope.name + (f" / {res.subpath}" if res.subpath else "")), F())
        ws.write_string(2, 0, "Groupes sources", lab)

        def lvl(g, label):
            if g is None:
                return f"{label} : aucun groupe"
            if g.error:
                return f"{label} : {g.stem} (fichier illisible)"
            return f"{label} : {g.stem} ({len(g.members)})"
        groups_txt = safe(f"{lvl(res.full, 'Modification')}   ·   {lvl(res.read, 'Lecture')}")
        ws.merge_range(2, 1, 2, c_com - 1, groups_txt, F(text_wrap=True, valign="top"))   # en-tête : hors zone de données
        if len(groups_txt) > 46:
            ws.set_row(2, 28)
        ws.write_string(3, 0, "Date de l'export", lab)
        ws.write_datetime(3, 1, dt.datetime.combine(scope.export_date, dt.time()),
                          F(num_format="dd/mm/yyyy", align="left"))

        input_fmt = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"],
                      text_wrap=True, valign="top")

        # ---- tableau
        sn = _sq(info.name)
        table_cols = [{"header": h, "header_format": self._hdr(input_col=(i in dec_cols or i == c_com))}
                      for i, h in enumerate(cols)]
        ws.add_table(_HEADER_ROW, 0, last, c_key, {"name": self._table_name("Membres"), "style": None,
                                                    "columns": table_cols, "autofilter": True})
        ws.set_row(_HEADER_ROW, 30)

        bd = {"border": 1, "border_color": c["border"]}
        txt = F(**bd)
        dec_fmt = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"], align="center")
        key_fmt = F(font_color="#7F7F7F", **bd)

        def put(r, col, value):
            if value:
                ws.write_string(r, col, safe(value), txt)
            else:
                ws.write_blank(r, col, None, txt)

        if n == 0:
            ws.write_string(first, 0, NO_MEMBER, F(italic=True, font_color="#595959", **bd))
            for col in range(1, c_key + 1):
                if col == c_com:
                    ws.write_blank(first, col, None, input_fmt)
                else:
                    ws.write_blank(first, col, None, txt)
        for i, p in enumerate(members):
            r = first + i
            put(r, 0, p.name)
            put(r, 1, p.sam)
            if c_dn is not None:
                put(r, c_dn, p.dn)
            ws.write_string(r, c_right, LEVEL_LABELS[p.level], txt)
            for dc in dec_cols:
                if native:
                    ws.insert_checkbox(r, dc, False, dec_fmt)
                else:
                    ws.write_blank(r, dc, None, dec_fmt)
            ws.write_blank(r, c_com, None, input_fmt)
            ws.write_string(r, c_key, p.key, key_fmt)

        # ---- commentaire général sous le tableau (zone de saisie ; seule fusion de la feuille, hors zone de données)
        cr = last + 2
        ws.write_string(cr, 0, "Commentaire général", F(bold=True, font_color=c["title"], valign="top"))
        ws.merge_range(cr, 1, cr, c_com, "", input_fmt)
        ws.set_row(cr, 60)
        info.comment_cell = xl_rowcol_to_cell(cr, 1)

        # ---- validation : liste Oui / Non
        if not native and n:
            ws.data_validation(first, c_dec, last, c_dec, {
                "validate": "list", "source": DROPDOWN_VALUES, "ignore_blank": True,
                "input_title": "Accès autorisé ?", "input_message": "Oui = accès justifié · Non = accès à retirer",
                "error_title": "Valeur non autorisée", "error_message": "Choisissez Oui ou Non dans la liste.",
                "error_type": "stop"})

        # ---- mise en forme conditionnelle (ordre = priorité) : vert / rouge / gris zébré
        if n:
            R = first + 1  # numéro de ligne Excel de la première ligne de données
            dx = self.F.dxf
            d1 = xl_col_to_name(dec_cols[0])
            if native:
                d2 = xl_col_to_name(dec_cols[1])
                ok, ko = f"=AND(${d1}{R}=TRUE,${d2}{R}<>TRUE)", f"=AND(${d2}{R}=TRUE,${d1}{R}<>TRUE)"
            else:
                ok, ko = f'=${d1}{R}="Oui"', f'=${d1}{R}="Non"'
            rules = [(ok, dx(bg_color=c["ok"])), (ko, dx(bg_color=c["ko"])),
                     ("=MOD(ROW(),2)=0", dx(bg_color=c["todo_even"])), ("=MOD(ROW(),2)=1", dx(bg_color=c["todo_odd"]))]
            for crit, fmt in rules:
                ws.conditional_format(first, 0, last, dec_cols[-1], {"type": "formula", "criteria": crit,
                                                                     "format": fmt, "stop_if_true": True})

        # ---- liens et progression (haut à droite)
        link = F(font_color="#0563C1", underline=1)
        rng = lambda col: f"'{sn}'!${xl_col_to_name(col)}${first + 1}:${xl_col_to_name(col)}${last + 1}"  # noqa: E731
        if n == 0:
            ws.write_string(0, c_com, "Progression : —", F(bold=True, font_color=c["title"]))
        else:
            ws.write_formula(0, c_com, '="Progression : "&(' + self._treated_formula(rng, dec_cols).lstrip("=") + f')&" / {n}"',
                             F(bold=True, font_color=c["title"], align="left"), f"Progression : 0 / {n}")
        ws.write_url(1, c_com, "internal:'Synthèse'!A1", link, "← Retour à la synthèse")
        ws.write_url(2, c_com, f"internal:'Synthèse'!A{self.add_title + 1}", link, "↓ Personnes à ajouter (onglet Synthèse)")
        ws.write_url(3, c_com, f"internal:'{sn}'!A{cr + 1}", link, "↓ Commentaire général (sous le tableau)")

        # ---- volets, impression, protection
        ws.freeze_panes(first, 0)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(_HEADER_ROW)
        ws.print_area(0, 0, cr, c_com)
        ws.set_margins(0.4, 0.4, 0.6, 0.6)
        ws.set_footer("&L&8&A&R&8Page &P / &N")
        self._protect(ws)

        info.data_hash = _digest(_rows_for_hash(scope, res))

    def _treated_formula(self, rng, dec_cols) -> str:
        """Formule : nombre de lignes traitées d'une feuille (Oui/Non, ou exactement une case cochée)."""
        if self.native:
            a, b = dec_cols
            return f"=SUMPRODUCT(--(({rng(a)}=TRUE)+({rng(b)}=TRUE)=1))"
        return f'=COUNTIF({rng(dec_cols[0])},"Oui")+COUNTIF({rng(dec_cols[0])},"Non")'

    # ------------------------------------------------------------------ Instructions (consignes + bloc de signature)
    def _instructions(self, ws):
        F, c, t = self.F, self.c, self.cfg["texts"]
        rv = self.cfg["review"]
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 6)
        ws.set_column(1, 1, 34)
        ws.set_column(2, 2, 26)      # C, D, E : 76 au total ; la date de validation y tient sur une ligne (jour | mois | année)
        ws.set_column(3, 4, 25)
        wrap = F(text_wrap=True, valign="top")
        w = 110
        bd = {"border": 1, "border_color": c["border"]}

        def put(r, text, fmt=wrap):
            ws.merge_range(r, 1, r, 4, safe(text), fmt)
            ws.set_row(r, max(16, 14 * math.ceil(len(text) / (w * 1.0)) + 3))

        values = {"label": self.book.label, "period": self.period, "contact": rv["contact"],
                  "deadline": rv["deadline"], "organisation": rv.get("organisation") or ""}
        r = 1
        ws.write_string(r, 1, t["instructions_title"].format(**values), F(bold=True, font_size=16, font_color=c["title"]))
        ws.set_row(r, 28)
        r += 1
        if values["organisation"]:
            ws.write_string(r, 1, values["organisation"], F(italic=True, font_color="#595959"))
            r += 1
        r += 1
        ws.write_string(r, 1, "Objectif", F(bold=True, font_color=c["title"], font_size=12)); r += 1
        put(r, " ".join(t["objective"].split())); r += 2
        ws.write_string(r, 1, "Comment répondre", F(bold=True, font_color=c["title"], font_size=12)); r += 1
        steps = t["steps_native"] if self.native else t["steps_dropdown"]
        for i, s in enumerate(steps, 1):
            ws.write_number(r, 0, i, F(align="right", valign="top", bold=True))
            put(r, s); r += 1
        r += 1
        ws.write_string(r, 1, "Date limite et contact", F(bold=True, font_color=c["title"], font_size=12)); r += 1
        put(r, t["return_line"].format(**values), F(text_wrap=True, valign="top", bold=True)); r += 1
        ws.write_string(r, 1, f"Contact : {values['contact']}", F()); r += 2

        # ---- validation par le propriétaire du groupe (signature électronique simple)
        ws.write_string(r, 1, "Validation et signature", F(bold=True, font_color=c["title"], font_size=12)); r += 1
        fields = ["Nom et prénom", "Fonction", "Date de validation (jour / mois / année)",
                  "Je certifie avoir examiné toutes les lignes et le récapitulatif", "Signature (saisissez votre nom)"]
        lab = F(bold=True, valign="vcenter", text_wrap=True, bg_color="#F2F2F2", **bd)
        inp = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"], valign="vcenter", text_wrap=True,
                align="left")
        sig_f = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"], italic=True,
                  font_size=self.cfg["excel"]["font_size"] + 4, valign="vcenter")
        chk_f = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"], align="left", valign="vcenter")
        top = r
        status_row = top + len(fields)
        d_, m_, y_ = (f"{col}{top + 3}" for col in "CDE")      # jour, mois, année : une seule ligne
        self.sig = {"name": f"C{top + 1}", "role": f"C{top + 2}", "day": d_, "month": m_, "year": y_,
                    "date": f"C{top + 3}:E{top + 3}", "attest": f"C{top + 4}", "signature": f"C{top + 5}",
                    "status": f"C{status_row + 1}"}
        s_ = self.sig
        for i, k in enumerate(fields):
            ws.write_string(top + i, 1, k, lab)
            if i == 2:    # date : trois listes déroulantes sur la même ligne
                for col in (2, 3, 4):
                    ws.write_blank(top + i, col, None, inp)
            elif i == 3:
                ws.merge_range(top + i, 2, top + i, 4, "", chk_f)
                ws.insert_checkbox(top + i, 2, False, chk_f)             # attestation : vraie case à cocher (TRUE / FALSE)
            else:
                ws.merge_range(top + i, 2, top + i, 4, "", sig_f if i == 4 else inp)
        ws.set_row(top + 3, 30)
        ws.set_row(top + 4, 36)
        # Excel n'a pas de calendrier intégré dans un .xlsx : jour / mois / année se choisissent dans trois listes.
        # Mois « 01 - Janvier » : le numéro est lu par LEFT(...,2), indépendant de la langue d'Excel.
        months = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre",
                  "Décembre"]
        y0 = self.now.year
        for col, vals, msg in ((2, [str(n) for n in range(1, 32)], "Jour : choisissez dans la liste."),
                               (3, [f"{n:02d} - {m}" for n, m in enumerate(months, 1)], "Mois : choisissez dans la liste."),
                               (4, [str(y0 + n) for n in range(3)], "Année : choisissez dans la liste.")):
            ws.data_validation(top + 2, col, top + 2, col, {
                "validate": "list", "source": vals, "ignore_blank": True,
                "input_title": "Date de validation", "input_message": msg,
                "error_title": "Valeur non autorisée", "error_message": "Choisissez une valeur dans la liste."})
        # date retenue : valable seulement si les trois champs sont remplis et forment un jour qui existe (pas de 31 février)
        d_v, m_v, y_v = f'VALUE({d_}&"")', f"VALUE(LEFT({m_},2))", f'VALUE({y_}&"")'
        date_ok = f"IFERROR(AND(COUNTA({d_},{m_},{y_})=3,DAY(DATE({y_v},{m_v},{d_v}))={d_v}),FALSE)"
        ws.conditional_format(top + 2, 2, top + 2, 4, {                  # jour impossible : les trois cases passent en rouge
            "type": "formula", "criteria": f"=AND(COUNTA({d_},{m_},{y_})=3,NOT({date_ok}))",
            "format": self.F.dxf(bg_color=c["ko"])})
        ws.write_string(status_row, 1, "Statut de la validation", F(bold=True, valign="vcenter", bg_color="#F2F2F2", **bd))
        stat = F(bold=True, valign="vcenter", **bd)
        ws.merge_range(status_row, 2, status_row, 4, "", stat)
        ws.write_formula(status_row, 2,
                         f'=IF(AND(LEN({s_["name"]})>0,LEN({s_["signature"]})>0,{date_ok},{s_["attest"]}=TRUE),"Signé","Non signé")',
                         stat, "Non signé")
        for text, color in (('"Signé"', c["ok"]), ('"Non signé"', c["todo_even"])):
            ws.conditional_format(status_row, 2, status_row, 4, {"type": "cell", "criteria": "==", "value": text,
                                                                  "format": self.F.dxf(bg_color=color)})
        r = status_row + 1
        put(r, " ".join(t["signature_hint"].split()), F(italic=True, font_color="#595959", text_wrap=True, valign="top"))
        r += 2

        ws.write_string(r, 1, "Légende des couleurs", F(bold=True, font_color=c["title"], font_size=12)); r += 1
        leg = t["legend"]
        items = [("ok", c["ok"]), ("ko", c["ko"]), ("todo", c["todo_even"]), ("input", c["input_bg"])]
        for key, color in items:
            extra = {"border_color": c["input_border"]} if key == "input" else {"border_color": c["border"]}
            ws.write_blank(r, 0, None, F(bg_color=color, border=1, **extra))
            ws.write_string(r, 1, leg[key], F(indent=1))
            r += 1
        r += 1
        put(r, t["protection_note"], F(text_wrap=True, valign="top", italic=True, font_color="#595959"))
        ws.set_portrait(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.set_margins(0.5, 0.5, 0.6, 0.6)
        self._protect(ws)

    # ------------------------------------------------------------------ En-tête (page de garde du document)
    def _cover(self, ws, infos: list[SheetInfo]):
        """Page de garde : version / indice / date, titre, classification, tableau « Mise à jour » (historique)."""
        F, c, t, rv = self.F, self.c, self.cfg["texts"], self.cfg["review"]
        ws.hide_gridlines(2)
        for col, wd in enumerate((3, 26, 14, 16, 46, 16, 16, 14)):
            ws.set_column(col, col, wd)
        white = {"border": 5, "border_color": "#FFFFFF"}                     # filets blancs épais entre les cellules
        scopes = self.book.scopes
        exports = " ; ".join(sorted({s.export_date.strftime("%d/%m/%Y") for s in scopes}))
        author = (rv.get("author") or "").strip() or "Shine (génération automatique)"
        today = self.now.replace(hour=0, minute=0, second=0, microsecond=0)

        # ---- haut de page : version, indice, date (à droite) ; titre au centre ; emplacement du logo en B2
        side_l = F(align="center", valign="vcenter")
        side_r = F(align="center", valign="vcenter")
        ws.write_string(0, 6, "Version", side_l); ws.write_string(0, 7, safe(str(rv.get("version") or "")), side_r)
        ws.write_string(1, 6, "Indice", F(align="center", valign="vcenter", bottom=2, bottom_color=c["cover_line"]))
        ws.write_string(1, 7, safe(str(rv.get("index") or "")), side_r)
        ws.write_string(2, 6, "Date", side_l)
        ws.write_datetime(2, 7, today, F(num_format="dd/mm/yyyy", align="center", valign="vcenter"))
        ws.set_row(1, 30)
        ws.merge_range(1, 2, 1, 5, "Revue des droits NTFS",
                       F(bold=True, font_size=18, align="center", valign="vcenter", bottom=2, bottom_color=c["cover_line"]))
        ws.merge_range(2, 2, 2, 5, f"{self.book.label} – {self.period}", F(italic=True, align="center", font_color="#595959"))

        # ---- classification du document : le niveau retenu est mis en évidence
        ws.merge_range(4, 1, 4, 3, "Classification du document", F(bold=True, font_size=14, valign="vcenter"))
        ws.set_row(4, 22)
        chosen = str(rv.get("classification") or "")
        for j, name in enumerate(CLASSIFICATIONS):
            on = name == chosen
            ws.write_string(5, 1 + j, name, F(align="center", valign="vcenter", text_wrap=True, bold=on,
                                              font_color="#FFFFFF" if on else "#000000",
                                              bg_color=c["cover_head"] if on else "#FFFFFF",
                                              border=1, border_color=c["cover_line"]))
        ws.set_row(5, 28)
        ws.merge_range(6, 1, 6, 7, safe(t["confidentiality_note"]), F(italic=True, font_color="#595959", font_size=9,
                                                                       text_wrap=True, valign="top"))
        ws.set_row(6, 18)

        # ---- tableau « Mise à jour » : 1re ligne écrite par Shine, lignes libres pour la suite (qui, quand, quoi, validation)
        ws.merge_range(8, 1, 8, 7, "Mise à jour", F(bold=True, font_size=14, font_color="#FFFFFF", bg_color=c["cover_head"],
                                                    align="center", valign="vcenter", **white))
        ws.set_row(8, 22)
        head = F(bold=True, font_size=12, font_color="#FFFFFF", bg_color=c["cover_head"], align="center", valign="vcenter",
                 text_wrap=True, **white)
        for j, name in enumerate(UPDATE_COLS):
            ws.write_string(9, 1 + j, name, head)
        ws.set_row(9, 32)
        body = F(bg_color=c["cover_body"], font_size=10, valign="vcenter", text_wrap=True, **white)
        body_date = F(bg_color=c["cover_body"], font_size=10, valign="vcenter", num_format="dd/mm/yyyy", align="center", **white)
        free = F(bg_color=c["cover_body"], font_size=10, valign="vcenter", text_wrap=True, locked=False, **white)
        free_date = F(bg_color=c["cover_body"], font_size=10, valign="vcenter", num_format="dd/mm/yyyy", align="center",
                      locked=False, **white)
        r1 = 10
        ws.write_string(r1, 1, safe(author), body)
        ws.write_datetime(r1, 2, today, body_date)
        ws.write_string(r1, 3, safe(str(rv.get("version") or "")), F(bg_color=c["cover_body"], font_size=10, valign="vcenter",
                                                                      align="center", **white))
        ws.write_string(r1, 4, f"Création du classeur à partir des exports du {exports} (Shine v{TOOL_VERSION} : mise en forme, sans analyse).", body)
        ws.write_blank(r1, 5, None, free)                                              # approbation : à saisir
        ws.write_formula(r1, 6, f"='{SHEET_INSTR}'!${self.sig['status'][0]}${self.sig['status'][1:]}", body, "Non signé")  # validation = signature
        ws.write_string(r1, 7, "Émis", body)
        ws.set_row(r1, 30)
        n_free = self.cfg["excel"]["history_rows"]
        for i in range(n_free):
            rr = r1 + 1 + i
            for j in range(7):
                ws.write_blank(rr, 1 + j, None, free_date if j == 1 else free)
            ws.set_row(rr, 20)
        if n_free:
            ws.data_validation(r1 + 1, 2, r1 + n_free, 2, self._date_dv("Date de la modification"))
        self.hist_first, self.hist_last = r1 + 1, r1 + n_free

        notice = str(rv.get("footer_notice") or "").strip()
        ws.set_portrait(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.set_margins(0.7, 0.7, 0.75, 0.75)
        ws.set_footer((f"&C&8&KFF0000 {notice.replace('&', '&&')} " if notice else "") + "&L&8&A&R&8Page &P / &N")
        self._protect(ws)

    # ------------------------------------------------------------------ Synthèse
    def _synthese(self, ws, infos: list[SheetInfo]):
        F, c = self.F, self.c
        ws.hide_gridlines(2)
        heads = ["Ressource", "Dossier", "Membres Full", "Membres Read", "Lignes à traiter",
                 "Traitées", "Reste", "Avancement"]
        widths = [max([len(i.res.name) for i in infos] + [24]) + 4,
                  max([len(i.scope.name) for i in infos] + [len(g["group"]) for g in self.groups] + [22]) + 2,
                  13, 13, 14, 11, 10, 12]
        for j, wd in enumerate(widths):
            ws.set_column(j, j, min(wd, 48))
        ws.write_string(0, 0, f"Synthèse – {self.book.label} – {self.period}", F(bold=True, font_size=15, font_color=c["title"]))
        ws.set_row(0, 26)
        hrow = SYN_HEAD_ROW
        first, last, _tr, _at = syn_layout(len(infos))
        total_lines = sum(i.n_rows for i in infos)
        ws.write_string(1, 0, "Progression globale", F(bold=True, font_color=c["title"]))
        ws.write_formula(1, 1, f'=IFERROR(SUM(F{first + 1}:F{last + 1})/SUM(E{first + 1}:E{last + 1}),"—")',
                         F(bold=True, num_format="0%", align="left"), 0 if total_lines else "—")
        ws.write_string(2, 0, "Une ligne = une ligne d'un export. « Traitées » = Oui ou Non (ou une case cochée) ; vide = non traité.",
                        F(italic=True, font_color="#595959"))

        ws.add_table(hrow, 0, last, len(heads) - 1, {
            "name": self._table_name("Synthese"), "style": None, "autofilter": True,
            "columns": [{"header": h, "header_format": self._hdr()} for h in heads]})
        ws.set_row(hrow, 32)
        bd = {"border": 1, "border_color": c["border"]}
        txt, num = F(**bd), F(align="center", **bd)
        link = F(font_color="#0563C1", underline=1, **bd)
        pct = F(num_format="0%", align="center", **bd)
        for i, info in enumerate(infos):
            r = first + i
            res = info.res
            sn = _sq(info.name)
            ws.write_url(r, 0, f"internal:'{sn}'!A1", link, safe(res.name)[:255], f"Aller à l'onglet {info.name}")
            ws.write_string(r, 1, safe(info.scope.name), txt)
            for j, lvl in ((2, FULL), (3, READ)):
                v = res.count(lvl)
                ws.write_string(r, j, "illisible", num) if v is None else ws.write_number(r, j, v, num)
            ws.write_number(r, 4, info.n_rows, num)
            f0, l0 = info.first + 1, info.last + 1
            rng = lambda col, s=sn, a=f0, b=l0: f"'{s}'!${xl_col_to_name(col)}${a}:${xl_col_to_name(col)}${b}"  # noqa: E731
            if info.n_rows == 0:
                ws.write_number(r, 5, 0, num)
            else:
                ws.write_formula(r, 5, self._treated_formula(rng, info.dec_cols), num, 0)
            R = r + 1
            ws.write_formula(r, 6, f"=E{R}-F{R}", num, info.n_rows)
            ws.write_formula(r, 7, f'=IF(E{R}=0,"—",F{R}/E{R})', pct, 0 if info.n_rows else "—")
        if not infos:
            ws.write_string(first, 0, "Aucune ressource", F(italic=True, **bd))
        # ligne de total : cellules ordinaires juste sous le tableau (SUBTOTAL respecte les filtres)
        tr, TR = last + 1, last + 2
        tb = {"top": 2, "bottom": 1, "left": 1, "right": 1, "border_color": c["border"], "bg_color": "#F2F2F2"}
        tot = F(bold=True, align="center", **tb)
        ws.write_string(tr, 0, "Total", F(bold=True, **tb))
        ws.write_blank(tr, 1, None, tot)
        cached = {2: sum(i.res.count(FULL) or 0 for i in infos), 3: sum(i.res.count(READ) or 0 for i in infos),
                  4: total_lines, 5: 0, 6: total_lines}
        for j in range(2, 7):
            col = xl_col_to_name(j)
            ws.write_formula(tr, j, f"=SUBTOTAL(109,{col}{first + 1}:{col}{last + 1})", tot, cached[j])
        ws.write_formula(tr, 7, f'=IF(E{TR}=0,"—",F{TR}/E{TR})', F(num_format="0%", align="center", bold=True, **tb),
                         0 if total_lines else "—")

        # ---- tableau « Personnes à ajouter » (sous le total) : une ligne = une personne et le groupe (fichier CSV) visé
        at = self.add_title
        ws.write_string(at, 0, "Personnes à ajouter", F(bold=True, font_size=self.cfg["excel"]["font_size"] + 2, font_color=c["title"]))
        ws.write_string(at + 1, 0, "Il manque quelqu'un ? Indiquez son nom (et son identifiant s'il est connu), le groupe de ressource "
                                   "concerné (à choisir dans la liste) et la justification. Le droit (Lecture ou Modification) "
                                   "se déduit du groupe choisi.", F(italic=True, font_color="#595959"))
        h = at + 2
        (g0, g1), (j0, j1) = ADD_GROUP_SPAN, ADD_JUSTIF_SPAN
        assert j1 == len(heads) - 1
        self.add_header_row = h
        ws.write_string(h, 0, ADD_COLS[0], self._hdr(input_col=True))
        ws.write_string(h, 1, ADD_COLS[1], self._hdr(input_col=True))
        ws.merge_range(h, g0, h, g1, ADD_COLS[2], self._hdr(input_col=True))           # groupe : C:E (valeur lue en C)
        ws.merge_range(h, j0, h, j1, ADD_COLS[3], self._hdr(input_col=True))           # justification : F:H (valeur lue en F)
        na = self.cfg["excel"]["add_rows"]
        self.add_first, self.add_last = h + 1, h + na
        add_fmt = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"], text_wrap=True)
        for r in range(h + 1, h + na + 1):
            ws.write_blank(r, 0, None, add_fmt)
            ws.write_blank(r, 1, None, add_fmt)
            ws.merge_range(r, g0, r, g1, "", add_fmt)
            ws.merge_range(r, j0, r, j1, "", add_fmt)
        if self.groups:
            col = xl_col_to_name(META_C2)
            ws.data_validation(h + 1, g0, h + na, g0, {
                "validate": "list", "source": f"='_meta'!${col}$2:${col}${len(self.groups) + 1}", "ignore_blank": True,
                "error_title": "Groupe non reconnu",
                "error_message": "Choisissez un groupe dans la liste. S'il n'y figure pas, laissez la cellule vide "
                                 "et précisez le groupe dans la justification."})
        ws.freeze_panes(first, 0)
        ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.repeat_rows(hrow)
        self._protect(ws)

    # ------------------------------------------------------------------ Récapitulatif (alimenté par formules via _calc)
    def _recap(self, ws):
        F, c = self.F, self.c
        ws.hide_gridlines(2)
        for col, wd in enumerate((16, 24, 30, 18, 18, 50, 34)):
            ws.set_column(col, col, wd)
        ws.set_column(7, 7, 8, None, {"hidden": True})               # H : n° de ligne dans _calc (aide au calcul)
        bd = {"border": 1, "border_color": c["border"]}
        ws.write_string(0, 0, f"Récapitulatif des demandes – {self.book.label} – {self.period}",
                        F(bold=True, font_size=15, font_color=c["title"]))
        ws.set_row(0, 26)
        ws.write_string(1, 0, "Tableau calculé automatiquement (ne pas modifier) : accès à retirer (lignes « Non » / « Retirer ») et "
                              "personnes à ajouter (tableau de l'onglet Synthèse). Il se met à jour à mesure que vous répondez.",
                        F(italic=True, font_color="#595959"))
        na, nr = self.n_add, self.n_rem
        lab = F(bold=True, font_color=c["title"])
        cnt = F(bold=True, align="left")
        ws.write_string(2, 0, "Ajouts demandés", lab)
        ws.write_formula(2, 1, f"='{SHEET_CALC}'!$I${na + 1}", cnt, 0)
        ws.write_string(3, 0, "Retraits demandés", lab)
        if nr:
            ws.write_formula(3, 1, f"='{SHEET_CALC}'!$B${nr + 1}", cnt, 0)
        else:
            ws.write_number(3, 1, 0, cnt)
        cap = na + self.cap_rem
        if self.cap_rem < nr:
            ws.write_formula(4, 0, f'=IF($B$3+$B$4>{cap},"⚠ "&($B$3+$B$4)&" demandes au total : seules les {cap} premières sont '
                                   f'affichées ici (paramètre excel.recap_max_removals).","")', F(bold=True, font_color="#C00000"), "")

        h, first = RECAP_HEAD_ROW, RECAP_HEAD_ROW + 1
        last = first + cap - 1
        self.recap_header_row, self.recap_first, self.recap_last = h, first, last
        ws.add_table(h, 0, last, len(RECAP_COLS) - 1, {
            "name": self._table_name("Recap"), "style": None, "autofilter": True,
            "columns": [{"header": k, "header_format": self._hdr()} for k in RECAP_COLS]})
        ws.set_row(h, 30)
        ws.write_string(h, 7, "n° _calc", F(font_color="#7F7F7F"))
        body = F(text_wrap=True, valign="top")
        calc = SHEET_CALC
        add_rng = lambda col: f"'{calc}'!${col}$2:${col}${na + 1}"           # noqa: E731
        rem_rng = lambda col: f"'{calc}'!${col}$2:${col}${nr + 1}"           # noqa: E731
        for r in range(first, last + 1):
            R = r + 1
            p = f"ROWS(H${first + 1}:H{R})"
            m_add = f"MATCH({p},{add_rng('I')},0)"
            m_rem = f"MATCH({p}-$B$3,{rem_rng('B')},0)"
            helper = f'=IF({p}<=$B$3,{m_add},IF({p}-$B$3<=$B$4,{m_rem},""))' if nr else f'=IF({p}<=$B$3,{m_add},"")'
            ws.write_formula(r, 7, helper, F(font_color="#7F7F7F"), "")
            ws.write_formula(r, 0, f'=IF($H{R}="","",IF({p}<=$B$3,"Ajout","Retrait"))', body, "")

            def pick(col, add_col, rem_col=None, const=None):
                a = f"INDEX({add_rng(add_col)},$H{R})"
                if not nr:
                    rem = '""'
                elif const is not None:
                    rem = f'"{const}"'
                else:
                    rem = f"INDEX({rem_rng(rem_col)},$H{R})"
                ws.write_formula(r, col, f'=IF($H{R}="","",IF($A{R}="Ajout",{a},{rem}))', body, "")
            pick(1, "O", "G")
            pick(2, "J", "C")
            pick(3, "K", "D")
            pick(4, "N", const=REMOVAL_RIGHT)
            pick(5, "M", "E")
            pick(6, "L", "F")
        # bordures et couleurs seulement sur les lignes remplies (une seule règle appliquée par cellule : compatible Excel et LibreOffice)
        edge = {"border": 1, "border_color": c["border"]}
        for value, color in (('"Ajout"', c["ok"]), ('"Retrait"', c["ko"])):
            ws.conditional_format(first, 0, last, 0, {"type": "cell", "criteria": "==", "value": value,
                                                      "format": self.F.dxf(bg_color=color, **edge), "stop_if_true": True})
        ws.conditional_format(first, 0, last, len(RECAP_COLS) - 1, {
            "type": "formula", "criteria": f'=$H{first + 1}<>""', "format": self.F.dxf(**edge)})
        ws.freeze_panes(first, 0)
        ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.repeat_rows(h)
        ws.print_area(0, 0, last, len(RECAP_COLS) - 1)
        self._protect(ws)

    # ------------------------------------------------------------------ _calc (masqué) : aide au calcul du récapitulatif
    def _calc_sheet(self, ws, infos: list[SheetInfo]):
        """Une ligne par ligne de ressource (retraits) et par ligne du tableau d'ajout.

        Colonne B / I : compteur cumulé des demandes ; le récapitulatif retrouve la n-ième demande par
        MATCH(n; compteur; 0), sans fonction volatile ni tableau dynamique (compatible toutes versions d'Excel).
        """
        F = self.F
        heads = ["feuille", "cumul retraits", "nom", "identifiant", "commentaire", "groupe", "propriétaire", "",
                 "cumul ajouts", "nom", "identifiant", "groupe", "justification", "droit", "propriétaire", "", "dates proposées"]
        for j, h in enumerate(heads):
            if h:
                ws.write_string(0, j, h, F(bold=True))
        day0 = dt.datetime.combine(self.now.date(), dt.time())
        datef = F(num_format="dd/mm/yyyy", align="left")
        for k in range(self.cfg["excel"]["date_choices_days"] + 1):
            # colonne Q : dates proposées dans les listes, à partir du jour d'OUVERTURE du classeur (TODAY recalculé à l'ouverture)
            ws.write_formula(1 + k, 16, "=TODAY()" + (f"+{k}" if k else ""), datef,
                             (day0 + dt.timedelta(days=k) - dt.datetime(1899, 12, 30)).days)
        groups = {(g["sheet"], g["level"]): g["group"] for g in self.groups}
        r = 1
        for info in infos:
            if not info.n_rows:
                continue
            sn = _sq(info.name)
            dec = info.dec_cols
            com = info.dec_cols[-1] + 1
            for k, p in enumerate(info.res.rows):
                R = info.first + 1 + k                                   # numéro de ligne Excel dans la feuille de ressource
                cell = lambda col: f"'{sn}'!${xl_col_to_name(col)}${R}"  # noqa: E731
                if self.native:
                    cond = f"AND({cell(dec[1])}=TRUE,{cell(dec[0])}<>TRUE)"
                else:
                    cond = f'{cell(dec[0])}="Non"'
                prev = f"B{r}+" if r > 1 else ""
                ws.write_string(r, 0, safe(info.name), F())
                ws.write_formula(r, 1, f"={prev}IF({cond},1,0)", F(), 0)
                ws.write_formula(r, 2, f'={cell(0)}&""', F(), "")
                ws.write_formula(r, 3, f'={cell(1)}&""', F(), "")
                ws.write_formula(r, 4, f'={cell(com)}&""', F(), "")
                ws.write_string(r, 5, safe(groups.get((info.name, p.level), "")), F())
                ws.write_string(r, 6, safe(info.scope.label), F())
                r += 1
        # ajouts : lignes du tableau « Personnes à ajouter » de la Synthèse
        single = self.book.scopes[0].label if len(self.book.scopes) == 1 else ""
        gcol = lambda k: xl_col_to_name(META_C2 + META_GROUPS.index(k))   # noqa: E731
        ng = len(self.groups)
        for i in range(self.n_add):
            r = 1 + i
            S = self.add_first + 1 + i                                    # ligne Excel dans la Synthèse
            syn = lambda col: f"'{SHEET_SYN}'!${col}${S}"                 # noqa: E731
            used = f"LEN(TRIM({syn('A')}&{syn('B')}&{syn('C')}&{syn('F')}))>0"
            prev = f"I{r}+" if i else ""
            ws.write_formula(r, 8, f"={prev}IF({used},1,0)", F(), 0)
            ws.write_formula(r, 9, f'={syn("A")}&""', F(), "")
            ws.write_formula(r, 10, f'={syn("B")}&""', F(), "")
            ws.write_formula(r, 11, f'={syn("C")}&""', F(), "")
            ws.write_formula(r, 12, f'={syn("F")}&""', F(), "")
            if ng:
                meta = lambda k: f"'_meta'!${gcol(k)}$2:${gcol(k)}${ng + 1}"   # noqa: E731
                m = f"MATCH(L{r + 1},{meta('group')},0)"
                ws.write_formula(r, 13, f'=IFERROR(IF(INDEX({meta("level")},{m})="{FULL}","{LEVEL_LABELS[FULL]}","{LEVEL_LABELS[READ]}"),"")', F(), "")
                ws.write_formula(r, 14, f'=IFERROR(INDEX({meta("owner")},{m}),"{_q(single)}")', F(), single)
            else:
                ws.write_string(r, 13, "", F())
                ws.write_string(r, 14, safe(single), F())
        self._protect(ws)

    # ------------------------------------------------------------------ _meta (masqué)
    def _meta_sheet(self, ws, infos: list[SheetInfo]):
        F = self.F
        folders = " | ".join(s.name for s in self.book.scopes)
        dates = sorted({s.export_date.isoformat() for s in self.book.scopes})
        kv = [
            ("schema_version", str(SCHEMA_VERSION)),
            ("tool_version", TOOL_VERSION),
            ("generated_at", self.now.replace(microsecond=0).isoformat()),
            ("period", self.period),
            ("workbook_label", self.book.label),
            ("folders", folders),
            ("export_dates", " | ".join(dates)),
            ("checkbox_mode", self.mode),
            ("decision_values", "Oui|Non" if not self.native else "Conserver=TRUE|Retirer=TRUE (vide = non traité)"),
            ("n_resources", str(len(infos))),
            ("n_rows", str(sum(i.n_rows for i in infos))),
            ("data_hash", self._data_hash()),
            ("add_sheet", "Synthèse"),
            ("add_header_row", str(self.add_header_row + 1)),
            ("add_first_row", str(self.add_first + 1)),
            ("add_last_row", str(self.add_last + 1)),
            ("add_col_name", "A"),
            ("add_col_sam", "B (optionnel)"),
            ("add_col_group", "C (fusion C:E ; liste MetaGroups.group)"),
            ("add_col_justification", "F (fusion F:H)"),
            ("cover_sheet", SHEET_COVER),
            ("signature_sheet", SHEET_INSTR),
            ("signature_cells", f"nom {self.sig['name']} · fonction {self.sig['role']} · date (jour, mois, année) {self.sig['date']} · "
                                      f"attestation {self.sig['attest']} · signature {self.sig['signature']} · statut {self.sig['status']}"),
            ("cover_history_rows", f"{self.hist_first + 1}-{self.hist_last + 1} (colonnes B:H : " + ", ".join(UPDATE_COLS) + ")"),
            ("classification", str(self.cfg["review"].get("classification") or "")),
            ("recap_sheet", SHEET_RECAP),
            ("recap_header_row", str(self.recap_header_row + 1)),
            ("recap_first_row", str(self.recap_first + 1)),
            ("recap_last_row", str(self.recap_last + 1)),
            ("recap_note", "feuille dérivée (formules, via la feuille masquée _calc) : ne pas relire pour l'ingestion"),
            ("row_key", "sha1(dossier|sous-dossier|ressource|niveau|SamAccountName|n° d'occurrence)[:16], casse ignorée"),
        ]
        # bloc 1 : clé / valeur (colonnes A:B)
        ws.add_table(0, 0, len(kv), 1, {"name": self._table_name("MetaKV"), "style": None,
                                         "columns": [{"header": "Clé"}, {"header": "Valeur"}]})
        for i, (k, v) in enumerate(kv, 1):
            ws.write_string(i, 0, k, F())
            ws.write_string(i, 1, safe(v), F())
        # bloc 2 : feuilles
        sh, c0 = META_SHEETS, META_C0
        ws.add_table(0, c0, max(len(infos), 1), c0 + len(sh) - 1, {
            "name": self._table_name("MetaSheets"), "style": None, "columns": [{"header": h} for h in sh]})
        for i, info in enumerate(infos, 1):
            dec = ",".join(xl_col_to_name(x) for x in info.dec_cols)
            vals = [info.name, info.scope.name, info.res.subpath, info.res.name, _group_name(info.res.full),
                    _group_name(info.res.read), info.header_row + 1, info.first + 1, info.last + 1, info.n_rows,
                    info.comment_cell, dec, xl_col_to_name(info.key_col), info.data_hash]
            for j, v in enumerate(vals):
                (ws.write_number if isinstance(v, int) else ws.write_string)(i, c0 + j, v if isinstance(v, int) else safe(v), F())
        # bloc 3 : lignes (clé stable → ligne initiale)
        c1, rh = META_C1, META_ROWS
        total = sum(i.n_rows for i in infos)
        ws.add_table(0, c1, max(total, 1), c1 + len(rh) - 1, {
            "name": self._table_name("MetaRows"), "style": None, "columns": [{"header": h} for h in rh]})
        r = 1
        for info in infos:
            for k, p in enumerate(info.res.rows):
                ws.write_string(r, c1, info.name, F())
                ws.write_string(r, c1 + 1, p.key, F())
                ws.write_string(r, c1 + 2, safe(p.sam), F())
                ws.write_string(r, c1 + 3, p.level, F())
                ws.write_number(r, c1 + 4, info.first + 1 + k, F())
                r += 1
        # bloc 4 : groupes (source de la liste déroulante « Groupe de ressource » de la Synthèse)
        c2, gh = META_C2, META_GROUPS
        ws.add_table(0, c2, max(len(self.groups), 1), c2 + len(gh) - 1, {
            "name": self._table_name("MetaGroups"), "style": None, "columns": [{"header": h} for h in gh]})
        for i, g in enumerate(self.groups, 1):
            for j, k in enumerate(gh):
                ws.write_string(i, c2 + j, safe(g[k]), F())
        ws.set_column(0, 0, 18); ws.set_column(1, 1, 60)
        self._protect(ws)


def build_book(path, book: Book, cfg: dict, mode: str, period: str, now: dt.datetime | None = None) -> BookReport:
    """Génère le classeur d'un `Book` et renvoie son compte rendu."""
    return BookBuilder(path, book, cfg, mode, period, now or dt.datetime.now()).build()
