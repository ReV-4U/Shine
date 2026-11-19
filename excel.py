# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Génération des classeurs Excel (XlsxWriter).

Structure d'un classeur :
  Instructions | Synthèse | <une feuille par ressource> | _meta (masqué)

Une ligne de feuille = une ligne d'un CSV (aucune fusion Full/Read, aucun dédoublonnage).

Contrat de relecture (phase 2, commande `ingest`) :
  * jamais de cellule fusionnée dans les zones de données (seul le commentaire général, en
    en-tête de feuille, est fusionné) ;
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

SCHEMA_VERSION = 3
TOOL_VERSION = "2.1.0"

LEVEL_LABELS = {FULL: "Modification", READ: "Lecture"}
ADD_COLS = ("Nom", "Identifiant (optionnel)", "Groupe de ressource", "Justification")   # tableau « Personnes à ajouter » (Synthèse)
ADD_GROUP_SPAN = (2, 4)      # colonnes C:E fusionnées par ligne (liste « Groupe de ressource »), valeur lue en C
ADD_JUSTIF_SPAN = (5, 8)     # colonnes F:I fusionnées par ligne, valeur lue en F
# colonnes de l'onglet _meta (contrat phase 2) ; les blocs sont placés côte à côte
META_SHEETS = ["sheet", "folder", "subpath", "resource", "group_full", "group_read", "header_row", "first_data_row",
               "last_data_row", "n_rows", "comment_cell", "decision_cols", "key_col", "data_hash"]
META_ROWS = ["sheet", "key", "sam", "level", "initial_row"]
META_GROUPS = ["group", "sheet", "folder", "subpath", "resource", "level", "file"]
META_C0 = 3
META_C1 = META_C0 + len(META_SHEETS) + 1
META_C2 = META_C1 + len(META_ROWS) + 1          # colonne de la liste « group » (source de la liste déroulante)
SYN_HEAD_ROW = 3                                 # en-tête (0-based) du tableau de la Synthèse


def syn_layout(n_resources: int) -> tuple[int, int, int, int]:
    """(première ligne de données, dernière, ligne de total, titre du tableau « Personnes à ajouter »), 0-based."""
    first = SYN_HEAD_ROW + 1
    last = SYN_HEAD_ROW + max(n_resources, 1)
    return first, last, last + 1, last + 4
DROPDOWN_VALUES = ["Oui", "Non", "À vérifier"]
NO_MEMBER = "Aucun membre"

_BAD_SHEET = re.compile(r"[\[\]:*?/\\]")
_RESERVED = {"instructions", "synthèse", "synthese", "_meta", "history"}
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
        ws_instr = wb.add_worksheet("Instructions")
        ws_syn = wb.add_worksheet("Synthèse")

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

        self._instructions(ws_instr)
        self._synthese(ws_syn, infos)
        ws_meta = wb.add_worksheet("_meta")
        self._meta_sheet(ws_meta, infos)
        ws_meta.hide()
        ws_instr.activate()
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
                            "resource": info.res.name, "level": lvl, "file": g.rel_path.rsplit("/", 1)[-1]})
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
        cols += ["Conserver", "Retirer"] if native else ["Accès légitime ?"]
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

        # commentaire général (zone de saisie ; seule fusion du classeur, hors zone de données)
        ws.write_string(4, 0, "Commentaire général", lab)
        input_fmt = F(locked=False, bg_color=c["input_bg"], border=1, border_color=c["input_border"],
                      text_wrap=True, valign="top")
        ws.merge_range(4, 1, 4, c_com, "", input_fmt)
        ws.set_row(4, 48)
        info.comment_cell = xl_rowcol_to_cell(4, 1)

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

        # ---- validation : liste Oui / Non / À vérifier
        if not native and n:
            ws.data_validation(first, c_dec, last, c_dec, {
                "validate": "list", "source": DROPDOWN_VALUES, "ignore_blank": True,
                "input_title": "Accès légitime ?", "input_message": "Oui = justifié · Non = à retirer · À vérifier",
                "error_title": "Valeur non autorisée", "error_message": "Choisissez Oui, Non ou À vérifier dans la liste.",
                "error_type": "stop"})

        # ---- mise en forme conditionnelle (ordre = priorité) : vert / rouge / jaune / gris zébré
        if n:
            R = first + 1  # numéro de ligne Excel de la première ligne de données
            dx = self.F.dxf
            d1 = xl_col_to_name(dec_cols[0])
            if native:
                d2 = xl_col_to_name(dec_cols[1])
                ok, ko, pend = (f"=AND(${d1}{R}=TRUE,${d2}{R}<>TRUE)", f"=AND(${d2}{R}=TRUE,${d1}{R}<>TRUE)",
                                f"=AND(${d1}{R}=TRUE,${d2}{R}=TRUE)")
            else:
                ok, ko, pend = f'=${d1}{R}="Oui"', f'=${d1}{R}="Non"', f'=${d1}{R}="À vérifier"'
            rules = [(ok, dx(bg_color=c["ok"])), (ko, dx(bg_color=c["ko"])), (pend, dx(bg_color=c["pending"])),
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

        # ---- volets, impression, protection
        ws.freeze_panes(first, 0)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(_HEADER_ROW)
        ws.print_area(0, 0, last, c_com)
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

    # ------------------------------------------------------------------ Instructions
    def _instructions(self, ws):
        F, c, t = self.F, self.c, self.cfg["texts"]
        rv = self.cfg["review"]
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 6)
        ws.set_column(1, 1, 110)
        wrap = F(text_wrap=True, valign="top")
        w = 110

        def put(r, text, fmt=wrap):
            ws.write_string(r, 1, safe(text), fmt)
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
        ws.write_string(r, 1, f"Date limite : {values['deadline']}   ·   Contact : {values['contact']}", F()); r += 2
        ws.write_string(r, 1, "Légende des couleurs", F(bold=True, font_color=c["title"], font_size=12)); r += 1
        leg = t["legend"]
        items = [("ok", c["ok"]), ("ko", c["ko"]), ("pending", c["pending"]), ("todo", c["todo_even"]), ("input", c["input_bg"])]
        for key, color in items:
            extra = {"border_color": c["input_border"]} if key == "input" else {"border_color": c["border"]}
            ws.write_blank(r, 0, None, F(bg_color=color, border=1, **extra))
            ws.write_string(r, 1, leg[key], F(indent=1))
            r += 1
        r += 1
        put(r, t["protection_note"], F(text_wrap=True, valign="top", italic=True, font_color="#595959"))
        ws.set_portrait(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.set_margins(0.5, 0.5, 0.6, 0.6)
        self._protect(ws)

    # ------------------------------------------------------------------ Synthèse
    def _synthese(self, ws, infos: list[SheetInfo]):
        F, c = self.F, self.c
        ws.hide_gridlines(2)
        heads = ["Ressource", "Dossier", "Membres Full", "Membres Read", "Lignes à traiter",
                 "Traitées", "À préciser", "Reste", "Avancement"]
        widths = [max([len(i.res.name) for i in infos] + [24]) + 4,
                  max([len(i.scope.name) for i in infos] + [len(g["group"]) for g in self.groups] + [22]) + 2,
                  13, 13, 14, 11, 11, 10, 12]
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
                ws.write_number(r, 6, 0, num)
            else:
                ws.write_formula(r, 5, self._treated_formula(rng, info.dec_cols), num, 0)
                if self.native:
                    a, b = info.dec_cols
                    ws.write_formula(r, 6, f"=SUMPRODUCT(--(({rng(a)}=TRUE)+({rng(b)}=TRUE)=2))", num, 0)
                else:
                    ws.write_formula(r, 6, f'=COUNTIF({rng(info.dec_cols[0])},"À vérifier")', num, 0)
            R = r + 1
            ws.write_formula(r, 7, f"=E{R}-F{R}-G{R}", num, info.n_rows)
            ws.write_formula(r, 8, f'=IF(E{R}=0,"—",F{R}/E{R})', pct, 0 if info.n_rows else "—")
        if not infos:
            ws.write_string(first, 0, "Aucune ressource", F(italic=True, **bd))
        # ligne de total : cellules ordinaires juste sous le tableau (SUBTOTAL respecte les filtres)
        tr, TR = last + 1, last + 2
        tb = {"top": 2, "bottom": 1, "left": 1, "right": 1, "border_color": c["border"], "bg_color": "#F2F2F2"}
        tot = F(bold=True, align="center", **tb)
        ws.write_string(tr, 0, "Total", F(bold=True, **tb))
        ws.write_blank(tr, 1, None, tot)
        cached = {2: sum(i.res.count(FULL) or 0 for i in infos), 3: sum(i.res.count(READ) or 0 for i in infos),
                  4: total_lines, 5: 0, 6: 0, 7: total_lines}
        for j in range(2, 8):
            col = xl_col_to_name(j)
            ws.write_formula(tr, j, f"=SUBTOTAL(109,{col}{first + 1}:{col}{last + 1})", tot, cached[j])
        ws.write_formula(tr, 8, f'=IF(E{TR}=0,"—",F{TR}/E{TR})', F(num_format="0%", align="center", bold=True, **tb),
                         0 if total_lines else "—")

        # ---- tableau « Personnes à ajouter » (sous le total) : une ligne = une personne et le groupe (fichier CSV) visé
        at = self.add_title
        ws.write_string(at, 0, "Personnes à ajouter", F(bold=True, font_size=self.cfg["excel"]["font_size"] + 2, font_color=c["title"]))
        ws.write_string(at + 1, 0, "Il manque quelqu'un ? Indiquez son nom (et son identifiant s'il est connu), le groupe de ressource "
                                   "concerné (à choisir dans la liste) et la justification.", F(italic=True, font_color="#595959"))
        h = at + 2
        (g0, g1), (j0, j1) = ADD_GROUP_SPAN, ADD_JUSTIF_SPAN
        assert j1 == len(heads) - 1
        self.add_header_row = h
        ws.write_string(h, 0, ADD_COLS[0], self._hdr(input_col=True))
        ws.write_string(h, 1, ADD_COLS[1], self._hdr(input_col=True))
        ws.merge_range(h, g0, h, g1, ADD_COLS[2], self._hdr(input_col=True))           # groupe : C:E (valeur lue en C)
        ws.merge_range(h, j0, h, j1, ADD_COLS[3], self._hdr(input_col=True))           # justification : F:I (valeur lue en F)
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
            ("decision_values", "Oui|Non|À vérifier" if not self.native else "Conserver=TRUE|Retirer=TRUE (vide = non traité)"),
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
            ("add_col_justification", "F (fusion F:I)"),
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
