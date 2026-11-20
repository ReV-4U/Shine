# Shine — mise en forme des exports de droits pour revue. Copyright (c) 2026 ReV. Licence non commerciale (voir LICENSE).
# Conçu par ReV, développé avec l'assistance de Claude (Anthropic).
"""Chargement et validation de la configuration (config.yaml)."""
from __future__ import annotations

import copy
import datetime as dt
import re
from pathlib import Path
from typing import Any

import yaml

DEFAULT_CONFIG_PATH = Path(__file__).with_name("config.yaml")
_HEX = re.compile(r"^#[0-9A-Fa-f]{6}$")


class ConfigError(ValueError):
    """Configuration invalide."""


def _merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def _regex(p: str, where: str) -> None:
    try:
        re.compile(p)
    except re.error as e:
        raise ConfigError(f"Expression régulière invalide dans {where} : {e}") from e


def validate(cfg: dict) -> None:
    _regex(cfg["input"]["folder_pattern"], "input.folder_pattern")
    if cfg["excel"]["checkbox_mode"] not in ("native", "dropdown"):
        raise ConfigError("excel.checkbox_mode doit valoir 'native' ou 'dropdown'")
    for key, low in (("add_rows", 1), ("recap_max_removals", 0), ("history_rows", 0), ("date_choices_days", 1)):
        v = cfg["excel"][key]
        if not isinstance(v, int) or isinstance(v, bool) or v < low:
            raise ConfigError(f"excel.{key} doit être un entier ≥ {low}")
    for name, col in cfg["excel"]["colors"].items():
        if not _HEX.match(str(col)):
            raise ConfigError(f"Couleur invalide excel.colors.{name} : {col!r} (attendu #RRGGBB)")
    pat = str(cfg["output"]["folder_pattern"])
    try:
        pat.format(label="x", date="x")
        dt.datetime(2026, 1, 2).strftime(str(cfg["output"]["date_format"]))
    except (KeyError, IndexError, ValueError) as e:
        raise ConfigError(f"output.folder_pattern / output.date_format invalide : {e}") from e
    if re.search(r'[\\/:*?"<>|]', pat.replace("{label}", "").replace("{date}", "")):
        raise ConfigError("output.folder_pattern contient un caractère interdit dans un nom de dossier")
    if cfg["review"].get("classification") not in ("1-Public", "2-Restreint", "3-Confidentiel"):
        raise ConfigError("review.classification doit valoir '1-Public', '2-Restreint' ou '3-Confidentiel'")
    p = cfg["review"].get("period")
    if p is not None and not re.fullmatch(r"\d{4}-\d{2}", str(p)):
        raise ConfigError("review.period doit être au format AAAA-MM")


def load_config(path: str | Path | None = None) -> dict[str, Any]:
    """Charge la configuration par défaut, fusionnée avec `path` s'il est donné."""
    cfg = yaml.safe_load(DEFAULT_CONFIG_PATH.read_text(encoding="utf-8"))
    if path:
        user = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        cfg = _merge(cfg, user)
    validate(cfg)
    return cfg
