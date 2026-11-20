#!/usr/bin/env python3
# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Shine — point d'entrée.

* Double-clic sur Shine.exe (ou `python shine.py` sans argument) : une fenêtre demande le dossier global
  (ex. « Liaison »), le traitement se lance, puis le dossier de sortie s'ouvre.
* Avec des arguments (`Shine.exe --input … --output …`) : mode ligne de commande de revue_droits.py.

Mise en forme uniquement ; traitement 100 % local ; aucune donnée personnelle affichée.
"""
from __future__ import annotations

import contextlib
import datetime as dt
import io
import os
import sys
import threading
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import revue_droits  # noqa: E402

__version__ = "2.0.0"
TITLE = "Shine"


def suggest_deadline(configured: str | None, today: dt.date | None = None, days: int = 14) -> str:
    """Date proposée dans la fenêtre : celle de la configuration si c'est une date valide, sinon aujourd'hui + 14 jours."""
    today = today or dt.date.today()
    try:
        return revue_droits.parse_deadline(configured or "")
    except ValueError:
        return (today + dt.timedelta(days=days)).strftime("%d/%m/%Y")


def check_deadline(text: str, today: dt.date | None = None) -> str:
    """Valide la saisie : date lisible et non passée. Renvoie JJ/MM/AAAA ou lève ValueError (message affichable)."""
    today = today or dt.date.today()
    d = revue_droits._parse_date(text)
    if d < today:
        raise ValueError("la date est déjà passée")
    return d.strftime("%d/%m/%Y")


def process_folder(folder: str, output: str | None = None, deadline: str | None = None) -> tuple[int, str, Path | None]:
    """Traite `folder` avec la configuration par défaut (+ config.yaml voisin de l'exécutable). Renvoie (code, résumé, dossier de sortie).

    Le résumé ne contient que des compteurs et des noms de dossiers/fichiers produits.
    """
    argv = ["--input", folder] + (["--output", output] if output else []) + (["--deadline", deadline] if deadline else [])
    args = revue_droits.parse_args(argv)
    buf_out, buf_err = io.StringIO(), io.StringIO()
    try:
        with contextlib.redirect_stdout(buf_out), contextlib.redirect_stderr(buf_err):
            rc = revue_droits.run(args)
    except OSError as e:
        return 1, (f"Erreur d'accès aux fichiers ({e.strerror or type(e).__name__}).\n"
                   "1) Un classeur du même jour est peut-être ouvert dans Excel : fermez-le.\n"
                   "2) Windows (Accès contrôlé aux dossiers, antivirus, OneDrive) bloque peut-être l'écriture : "
                   "autorisez Shine.exe ou changez de dossier de sortie."), None
    except Exception as e:  # noqa: BLE001 - la fenêtre doit toujours pouvoir afficher quelque chose
        tb = traceback.extract_tb(e.__traceback__)
        where = f", {Path(tb[-1].filename).name}:{tb[-1].lineno}" if tb else ""          # lieu du code, jamais de données
        return 1, (f"Erreur inattendue ({type(e).__name__}{where}). Le traitement a été interrompu.\n"
                   "Relancez Shine ; si l'erreur persiste, transmettez ce message tel quel au développeur."), None
    text = (buf_err.getvalue() if rc else buf_out.getvalue()).strip()
    if rc == 0 and "=== Résumé ===" in text:
        text = text.split("=== Résumé ===", 1)[1].strip()
    return rc, text or "Aucune information.", getattr(args, "resolved_output", None)


def open_folder(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')  # noqa: S605
        else:
            os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')  # noqa: S605
    except OSError:
        pass


def gui() -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(parent=root, title="Shine — choisissez le dossier global (ex. : Liaison)",
                                     mustexist=True)
    if not folder:
        root.destroy()
        return 0

    # date limite : saisie par l'utilisateur au moment de compiler (proposition pré-remplie)
    try:
        import config as _config
        configured = _config.load_config(revue_droits.default_user_config())["review"].get("deadline")
    except Exception:  # noqa: BLE001
        configured = None
    deadline = None
    proposal = suggest_deadline(configured)
    while deadline is None:
        answer = simpledialog.askstring(
            TITLE, "Date limite de retour des classeurs par les ayants droit\n(format JJ/MM/AAAA) :",
            initialvalue=proposal, parent=root)
        if answer is None:                       # Annuler : on n'écrit rien
            root.destroy()
            return 0
        try:
            deadline = check_deadline(answer)
        except ValueError as e:
            proposal = answer
            messagebox.showwarning(TITLE, f"Date non valide ({e}).\nExemple : 30/11/2026", parent=root)

    # petite fenêtre d'attente pendant que le traitement tourne dans un thread
    root.deiconify()
    root.title(TITLE)
    root.geometry("380x90")
    root.resizable(False, False)
    tk.Label(root, text="Mise en forme en cours…\nMerci de patienter.", pady=22).pack()
    result: dict = {}
    t = threading.Thread(target=lambda: result.update(r=process_folder(folder, deadline=deadline)), daemon=True)
    t.start()

    def poll():
        if t.is_alive():
            root.after(150, poll)
        else:
            root.withdraw()
            rc, text, out = result["r"]
            if rc == 0:
                messagebox.showinfo(TITLE, f"Terminé.\n\n{text}", parent=root)
                if out:
                    open_folder(out)
            else:
                messagebox.showerror(TITLE, f"Le traitement n'a pas abouti.\n\n{text}", parent=root)
            root.destroy()
    root.after(150, poll)
    root.mainloop()
    return 0 if result.get("r", (1,))[0] == 0 else 1


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    return revue_droits.main(argv) if argv else gui()


if __name__ == "__main__":
    sys.exit(main())
