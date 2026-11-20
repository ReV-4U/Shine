# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Extraction du récapitulatif : lit un classeur renvoyé par un ayant droit et en tire un fichier à part.

Le fichier produit contient le bloc de validation (nom, fonction, date, attestation, signature, statut), les
compteurs et le tableau de l'onglet « Récapitulatif », en valeurs figées (aucune formule, aucun lien externe).

Les valeurs sont celles enregistrées par Excel dans le classeur renvoyé : il doit donc avoir été enregistré depuis
Excel (ou LibreOffice) après les réponses. Aucune donnée n'est analysée ni interprétée : copie fidèle.
"""
from __future__ import annotations

import datetime as dt
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import xlsxwriter

from excel import SHEET_RECAP, TOOL_VERSION, safe

# « nom C10 · fonction C11 · date (jour, mois, année) C12:E12 · attestation C13 · signature C14 · statut C16 »
_SIG_RE = re.compile(r"nom (\w+) · fonction (\w+) · date \(jour, mois, année\) (\w+):(\w+) · "
                     r"attestation (\w+) · signature (\w+) · statut (\w+)")
_COLORS = {"head": "#1F3864", "border": "#BFBFBF", "ok": "#C6EFCE", "ko": "#FFC7CE", "todo": "#FFEB9C", "label": "#F2F2F2"}


class ExtractError(Exception):
    """Classeur illisible ou non reconnu : le message est affichable tel quel (aucune donnée personnelle)."""


@dataclass
class Extraction:
    path: Path
    n_add: int = 0
    n_rem: int = 0
    n_rows: int = 0
    status: str = ""
    warnings: list[str] = field(default_factory=list)


def _meta(wb) -> dict[str, str]:
    if "_meta" not in wb.sheetnames:
        raise ExtractError("Ce fichier n'est pas un classeur produit par Shine (onglet technique « _meta » absent).")
    ws, out = wb["_meta"], {}
    for key, val in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        if key is None:
            break
        out[str(key)] = "" if val is None else str(val)
    return out


def _text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Oui" if v else "Non"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _date_text(day, month, year) -> tuple[str, bool]:
    """Date JJ/MM/AAAA à partir des trois listes ; (texte, valide). Le mois est de la forme « 11 - Novembre »."""
    parts = [_text(day).strip(), _text(month).strip(), _text(year).strip()]
    if not all(parts):
        return " / ".join(p or "…" for p in parts), False
    try:
        return dt.date(int(parts[2]), int(parts[1][:2]), int(parts[0])).strftime("%d/%m/%Y"), True
    except ValueError:
        return " / ".join(parts), False


def read_recap(src: Path) -> dict:
    """Lit le classeur renvoyé et renvoie les valeurs à recopier (voir `extract_recap`)."""
    try:
        wb = openpyxl.load_workbook(src, data_only=True)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, openpyxl.utils.exceptions.InvalidFileException) as e:
        raise ExtractError(f"Fichier illisible ({type(e).__name__}) : est-ce bien un classeur .xlsx ?") from e
    meta = _meta(wb)
    sheet = meta.get("recap_sheet") or SHEET_RECAP
    if sheet not in wb.sheetnames:
        raise ExtractError("Onglet « Récapitulatif » introuvable : classeur modifié ou trop ancien.")
    m = _SIG_RE.fullmatch(meta.get("signature_cells", ""))
    if not m or meta.get("signature_sheet") not in wb.sheetnames:
        raise ExtractError("Bloc de signature non reconnu : classeur généré par une version trop ancienne de Shine. "
                           "Régénérez-le avec la version actuelle.")
    name, role, d1, d3, attest, signature, status = m.groups()
    day, month, year = (f"{c}{d1[1:]}" for c in "CDE")
    si = wb[meta["signature_sheet"]]
    day_v, month_v, year_v = (si[c].value for c in (day, month, year))
    date_txt, date_ok = _date_text(day_v, month_v, year_v)

    rc = wb[sheet]
    first, last, head = (int(meta[k]) for k in ("recap_first_row", "recap_last_row", "recap_header_row"))
    header = [_text(c) for c in next(rc.iter_rows(min_row=head, max_row=head, max_col=7, values_only=True))]
    rows = [[_text(v) for v in r] for r in rc.iter_rows(min_row=first, max_row=last, max_col=7, values_only=True)]
    rows = [r for r in rows if r[0]]                                    # lignes vides = formules sans demande
    n_add, n_rem = (rc[c].value for c in ("B3", "B4"))
    return {
        "label": meta.get("workbook_label", ""), "period": meta.get("period", ""), "folders": meta.get("folders", ""),
        "export_dates": meta.get("export_dates", ""), "classification": meta.get("classification", ""),
        "header": header, "rows": rows, "n_add": n_add, "n_rem": n_rem,
        "sig": {"name": _text(si[name].value), "role": _text(si[role].value), "date": date_txt, "date_ok": date_ok,
                "attest": bool(si[attest].value is True), "signature": _text(si[signature].value),
                "status": _text(si[status].value)},
    }


def _write(path: Path, src: Path, data: dict, now: dt.datetime) -> None:
    wb = xlsxwriter.Workbook(str(path))
    wb.set_properties({"title": "Récapitulatif des demandes", "comments": f"Extrait par Shine v{TOOL_VERSION} (valeurs figées)"})
    ws = wb.add_worksheet("Récapitulatif")
    F = lambda **kw: wb.add_format({"font_name": "Calibri", "font_size": 11, "valign": "top", **kw})   # noqa: E731
    c, bd = _COLORS, {"border": 1, "border_color": _COLORS["border"]}
    ws.hide_gridlines(2)
    for col, wd in enumerate((24, 24, 30, 18, 18, 50, 34)):
        ws.set_column(col, col, wd)
    ws.write_string(0, 0, safe(f"Récapitulatif des demandes – {data['label']} – {data['period']}"),
                    F(bold=True, font_size=15, font_color=c["head"]))
    ws.set_row(0, 26)
    ws.write_string(1, 0, safe(f"Extrait de « {src.name} » le {now:%d/%m/%Y} par Shine v{TOOL_VERSION} : valeurs figées, "
                               "copie fidèle de l'onglet Récapitulatif et du bloc de signature."),
                    F(italic=True, font_color="#595959"))
    sig, lab = data["sig"], F(bold=True, bg_color=c["label"], **bd)
    lines = [("Statut de la validation", sig["status"] or "Non signé"), ("Nom et prénom", sig["name"]),
             ("Fonction", sig["role"]), ("Date de validation", sig["date"] + ("" if sig["date_ok"] else "  (incomplète ou invalide)")),
             ("Attestation d'examen", "Oui" if sig["attest"] else "Non"), ("Signature", sig["signature"])]
    r = 3
    for k, v in lines:
        ws.write_string(r, 0, k, lab)
        fmt = F(bold=k.startswith("Statut"), **bd)
        ws.merge_range(r, 1, r, 2, safe(v), fmt)
        r += 1
    ws.conditional_format(3, 1, 3, 2, {"type": "cell", "criteria": "==", "value": '"Signé"', "format": wb.add_format({"bg_color": c["ok"]})})
    ws.conditional_format(3, 1, 3, 2, {"type": "cell", "criteria": "==", "value": '"Non signé"', "format": wb.add_format({"bg_color": c["todo"]})})
    r += 1
    n_add, n_rem = int(data["n_add"] or 0), int(data["n_rem"] or 0)
    ws.write_string(r, 0, "Ajouts demandés", F(bold=True, font_color=c["head"]))
    ws.write_number(r, 1, n_add, F(bold=True, align="left"))
    ws.write_string(r + 1, 0, "Retraits demandés", F(bold=True, font_color=c["head"]))
    ws.write_number(r + 1, 1, n_rem, F(bold=True, align="left"))
    h = r + 3
    head = F(bold=True, font_color="#FFFFFF", bg_color=c["head"], text_wrap=True, valign="vcenter", **bd)
    ncol = len(data["header"])
    rows = data["rows"]
    ws.add_table(h, 0, h + max(len(rows), 1), ncol - 1, {
        "name": "Recap", "style": None, "autofilter": True,
        "columns": [{"header": safe(k), "header_format": head} for k in data["header"]]})
    ws.set_row(h, 30)
    body = F(text_wrap=True, **bd)
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            ws.write_string(h + i, j, safe(v), body)
    for value, color in (('"Ajout"', c["ok"]), ('"Retrait"', c["ko"])):
        ws.conditional_format(h + 1, 0, h + max(len(rows), 1), 0, {
            "type": "cell", "criteria": "==", "value": value, "format": wb.add_format({"bg_color": color})})
    ws.freeze_panes(h + 1, 0)
    ws.set_landscape(); ws.set_paper(9); ws.fit_to_pages(1, 0); ws.repeat_rows(h)
    ws.set_footer("&L&8&A&R&8Page &P / &N")
    wb.close()


def extract_recap(src, out_dir=None, now: dt.datetime | None = None) -> Extraction:
    """Crée `Recap_<nom du classeur>.xlsx` (à côté du classeur, ou dans `out_dir`) et renvoie un compte rendu sans données."""
    src, now = Path(src), now or dt.datetime.now()
    data = read_recap(src)
    out = Path(out_dir) if out_dir else src.parent
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"Recap_{src.stem}.xlsx"
    if path.resolve() == src.resolve():
        raise ExtractError("Le fichier d'entrée est déjà un récapitulatif extrait.")
    _write(path, src, data, now)
    res = Extraction(path=path, n_add=int(data["n_add"] or 0), n_rem=int(data["n_rem"] or 0), n_rows=len(data["rows"]),
                     status=data["sig"]["status"] or "Non signé")
    if data["n_add"] is None or data["n_rem"] is None:
        res.warnings.append("Aucune valeur enregistrée dans le récapitulatif : ouvrez le classeur dans Excel, enregistrez-le, "
                            "puis relancez l'extraction.")
    elif res.n_rows < res.n_add + res.n_rem:
        res.warnings.append(f"{res.n_add + res.n_rem} demandes au total, mais {res.n_rows} lignes affichées dans l'onglet "
                            "Récapitulatif (limite excel.recap_max_removals).")
    if res.status != "Signé":
        res.warnings.append("Le classeur n'est pas signé (nom, date, case à cocher et signature requis).")
    return res


def collect(paths) -> tuple[list[Path], bool]:
    """Fichiers à traiter. Un dossier est parcouru récursivement (sans `~$…` ni `Recap_…`) ; renvoie (fichiers, via_dossier)."""
    out, folder = [], False
    for p in map(Path, paths):
        if p.is_dir():
            folder = True
            out += sorted(f for f in p.rglob("*.xlsx") if f.is_file() and not f.name.startswith(("~$", "Recap_")))
        else:
            out.append(p)
    return out, folder


def _is_shine(path: Path) -> bool:
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        try:
            return "_meta" in wb.sheetnames
        finally:
            wb.close()
    except Exception:  # noqa: BLE001 - fichier illisible : simplement ignoré lors d'un parcours de dossier
        return False


def main(files, output: str | None = None, now: dt.datetime | None = None) -> tuple[int, str]:
    """Traite des classeurs ou des dossiers ; renvoie (code, résumé). Le résumé ne contient que des compteurs et des noms de fichiers.

    Dans un dossier, les fichiers qui ne sont pas des classeurs Shine sont ignorés (comptés) ; un fichier donné
    explicitement doit en être un, sinon l'erreur est signalée.
    """
    lines, bad, skipped = [], 0, 0
    found, via_folder = collect(files)
    explicit = {Path(f).resolve() for f in files if not Path(f).is_dir()}
    for f in found:
        name = Path(f).name
        if via_folder and Path(f).resolve() not in explicit and not _is_shine(Path(f)):
            skipped += 1
            continue
        try:
            e = extract_recap(f, output, now)
        except ExtractError as err:
            bad += 1
            lines.append(f"✗ {name} : {err}")
            continue
        except OSError as err:
            bad += 1
            lines.append(f"✗ {name} : écriture impossible ({err.strerror or type(err).__name__}) – "
                         "un fichier Recap du même nom est peut-être ouvert dans Excel.")
            continue
        lines.append(f"✓ {name} → {e.path.name} : {e.n_add} ajout(s), {e.n_rem} retrait(s), statut « {e.status} »")
        lines += [f"    ! {w}" for w in e.warnings]
    if skipped:
        lines.append(f"{skipped} fichier(s) Excel ignoré(s) (pas des classeurs Shine).")
    return (1 if bad else 0), "\n".join(lines) or "Aucun classeur Shine trouvé."
