"""Helpers puros para autoselección de templates."""

import os
import re
import unicodedata


def normalize_match_text(text):
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def tokenize_match_text(text, ignore=None):
    ignore = ignore or set()
    return [tok for tok in normalize_match_text(text).split() if len(tok) > 2 and tok not in ignore]


def score_template_candidate(producto, cliente, path, ignore=None):
    """Asigna puntaje a un template según coincidencia de producto y cliente."""
    ignore = ignore or set()
    filename = os.path.basename(path)
    folder = os.path.dirname(path)
    fname_n = normalize_match_text(filename)
    folder_n = normalize_match_text(folder)
    prod_n = normalize_match_text(producto)

    score = 0
    if prod_n and prod_n in fname_n:
        score += 12
    elif prod_n and prod_n in folder_n:
        score += 7

    for tok in tokenize_match_text(producto, ignore):
        if tok in fname_n:
            score += 4
        elif tok in folder_n:
            score += 2

    for tok in tokenize_match_text(cliente, ignore):
        if tok in folder_n:
            score += 3
        elif tok in fname_n:
            score += 1

    if "template" in fname_n:
        score += 1
    if "mix" in fname_n and "mix" not in prod_n and "blend" not in prod_n:
        score -= 2

    return score
