"""
Category classifier and name standardizer.

Usage:
  Train:   python classifier.py train  "raw name"  "Category Name"
  Classify: python classifier.py classify "raw name"
  List:    python classifier.py list
  Reset:   python classifier.py reset "Category Name"   (removes a category)
"""

import json
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

DB_PATH = Path(__file__).parent / "categories.json"


# ── persistence ──────────────────────────────────────────────────────────────

def load() -> dict:
    if DB_PATH.exists():
        return json.loads(DB_PATH.read_text())
    return {"categories": {}}


def save(db: dict) -> None:
    DB_PATH.write_text(json.dumps(db, indent=2))


# ── text helpers ─────────────────────────────────────────────────────────────

def normalize(text: str) -> str:
    """Lowercase, strip punctuation/extra spaces for comparison."""
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize(a), normalize(b)).ratio()


# ── core logic ────────────────────────────────────────────────────────────────

def train(raw_name: str, category: str) -> None:
    """Register a raw name as belonging to a category."""
    db = load()
    cats = db["categories"]

    # Canonical category name: title-case the provided value
    canon_cat = category.strip().title()

    if canon_cat not in cats:
        # First entry becomes the canonical display name for the category
        cats[canon_cat] = {"canonical": canon_cat, "examples": []}

    norm = normalize(raw_name)
    examples = cats[canon_cat]["examples"]

    if norm not in [normalize(e) for e in examples]:
        examples.append(raw_name.strip())
        print(f'Trained: "{raw_name.strip()}" -> [{canon_cat}]')
    else:
        print(f'Already known: "{raw_name.strip()}" is [{canon_cat}]')

    save(db)


def classify(raw_name: str, threshold: float = 0.55) -> tuple[str | None, str, float]:
    """
    Return (category, standardized_name, confidence).
    category is None when confidence is below threshold.
    """
    db = load()
    cats = db["categories"]

    if not cats:
        return None, raw_name.strip(), 0.0

    best_cat = None
    best_score = 0.0
    best_example = raw_name.strip()

    for cat_name, cat_data in cats.items():
        for example in cat_data["examples"]:
            score = similarity(raw_name, example)
            if score > best_score:
                best_score = score
                best_cat = cat_name
                best_example = example

    if best_score < threshold:
        return None, raw_name.strip(), best_score

    # Standardized name = the first registered example for the category
    # (acts as the canonical form for that category's entries)
    canonical_example = cats[best_cat]["examples"][0]
    standardized = _standardize(raw_name, canonical_example, best_cat)

    return best_cat, standardized, best_score


def _standardize(raw: str, canonical_example: str, category: str) -> str:
    """
    Produce a standardized form of raw based on its category's canonical example.
    Strategy: preserve the raw token structure but apply the casing/spacing
    pattern of the canonical example.
    """
    # Simple rule: title-case words, strip extra whitespace
    standardized = " ".join(w.capitalize() for w in raw.strip().split())
    return standardized


def list_categories() -> None:
    db = load()
    cats = db["categories"]
    if not cats:
        print("No categories trained yet.")
        return
    for cat_name, cat_data in cats.items():
        examples = cat_data["examples"]
        print(f"\n[{cat_name}]  ({len(examples)} example{'s' if len(examples) != 1 else ''})")
        for ex in examples:
            print(f"  - {ex}")


def reset_category(category: str) -> None:
    db = load()
    canon = category.strip().title()
    if canon in db["categories"]:
        del db["categories"][canon]
        save(db)
        print(f"Removed category [{canon}].")
    else:
        print(f"Category [{canon}] not found.")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    args = sys.argv[1:]

    if not args:
        print(__doc__)
        return

    cmd = args[0].lower()

    if cmd == "train":
        if len(args) < 3:
            print("Usage: python classifier.py train <name> <category>")
            sys.exit(1)
        train(args[1], args[2])

    elif cmd == "classify":
        if len(args) < 2:
            print("Usage: python classifier.py classify <name>")
            sys.exit(1)
        raw = " ".join(args[1:])
        category, standardized, confidence = classify(raw)
        if category is None:
            print(f'Input:       "{raw}"')
            print(f"Category:    UNKNOWN  (best confidence {confidence:.0%})")
            print(f"Standardized: {standardized}")
        else:
            print(f'Input:        "{raw}"')
            print(f"Category:     [{category}]  (confidence {confidence:.0%})")
            print(f"Standardized: {standardized}")

    elif cmd == "list":
        list_categories()

    elif cmd == "reset":
        if len(args) < 2:
            print("Usage: python classifier.py reset <category>")
            sys.exit(1)
        reset_category(" ".join(args[1:]))

    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
