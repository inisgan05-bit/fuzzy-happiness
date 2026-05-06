"""
Category classifier and name standardizer.

Commands:
  train     python classifier.py train "raw name" "Standardized Name"
  classify  python classifier.py classify "raw name"
  batch     python classifier.py batch input.txt          (one name per line)
  list      python classifier.py list
  reset     python classifier.py reset "Category Name"
  stats     python classifier.py stats
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
    text = text.lower().strip()
    # Remove leading numbering like "2. " or "13. "
    text = re.sub(r"^\d+\.\s*", "", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize(a), normalize(b)).ratio()


def keyword_boost(raw: str, cat_name: str) -> float:
    """Small boost when key tokens from category name appear in the raw name."""
    raw_n = normalize(raw)
    cat_n = normalize(cat_name)
    tokens = [t for t in cat_n.split() if len(t) > 3]
    if not tokens:
        return 0.0
    hits = sum(1 for t in tokens if t in raw_n)
    return 0.05 * (hits / len(tokens))


# ── core logic ────────────────────────────────────────────────────────────────

def train(raw_name: str, category: str, db: dict | None = None, silent: bool = False) -> dict:
    """Register a raw name as belonging to a category. Returns the db."""
    own_db = db is None
    if own_db:
        db = load()

    cats = db["categories"]
    canon_cat = category.strip()

    if canon_cat not in cats:
        cats[canon_cat] = {"canonical": canon_cat, "examples": []}

    norm_raw = normalize(raw_name)
    examples = cats[canon_cat]["examples"]
    norm_existing = [normalize(e) for e in examples]

    if norm_raw and norm_raw not in norm_existing:
        examples.append(raw_name.strip())
        if not silent:
            print(f'  Trained: "{raw_name.strip()}" -> [{canon_cat}]')

    if own_db:
        save(db)
    return db


def bulk_train(pairs: list[tuple[str, str]], silent: bool = False) -> None:
    """Load many (raw_name, category) pairs efficiently."""
    db = load()
    count = 0
    for raw, cat in pairs:
        raw = raw.strip()
        cat = cat.strip()
        if not raw or not cat or raw in ("??", "?"):
            continue
        before = len(db["categories"].get(cat, {}).get("examples", []))
        db = train(raw, cat, db=db, silent=silent)
        after = len(db["categories"].get(cat, {}).get("examples", []))
        if after > before:
            count += 1
    save(db)
    print(f"Bulk import complete: {count} new entries added across {len(db['categories'])} categories.")


def classify(raw_name: str, threshold: float = 0.50) -> tuple[str | None, str, float]:
    """
    Return (standardized_category, standardized_name, confidence).
    standardized_category is None when confidence is below threshold.
    The standardized_name IS the category name.
    """
    db = load()
    cats = db["categories"]

    if not cats:
        return None, raw_name.strip(), 0.0

    raw_name = raw_name.strip()
    best_cat = None
    best_score = 0.0

    for cat_name, cat_data in cats.items():
        for example in cat_data["examples"]:
            score = similarity(raw_name, example) + keyword_boost(raw_name, cat_name)
            if score > best_score:
                best_score = score
                best_cat = cat_name

    if best_score < threshold:
        return None, raw_name, best_score

    return best_cat, best_cat, best_score


def batch_classify(names: list[str], threshold: float = 0.50) -> list[tuple[str, str | None, str, float]]:
    """Classify a list of names. Returns list of (raw, category, standardized, confidence)."""
    db = load()
    cats = db["categories"]
    results = []

    for raw_name in names:
        raw_name = raw_name.strip()
        if not raw_name:
            continue

        best_cat = None
        best_score = 0.0

        for cat_name, cat_data in cats.items():
            for example in cat_data["examples"]:
                score = similarity(raw_name, example) + keyword_boost(raw_name, cat_name)
                if score > best_score:
                    best_score = score
                    best_cat = cat_name

        if best_score < threshold:
            results.append((raw_name, None, raw_name, best_score))
        else:
            results.append((raw_name, best_cat, best_cat, best_score))

    return results


def list_categories() -> None:
    db = load()
    cats = db["categories"]
    if not cats:
        print("No categories trained yet.")
        return
    for cat_name in sorted(cats):
        examples = cats[cat_name]["examples"]
        print(f"\n[{cat_name}]  ({len(examples)} example{'s' if len(examples) != 1 else ''})")
        for ex in examples[:5]:
            print(f"  - {ex}")
        if len(examples) > 5:
            print(f"  ... and {len(examples) - 5} more")


def reset_category(category: str) -> None:
    db = load()
    if category in db["categories"]:
        del db["categories"][category]
        save(db)
        print(f"Removed category [{category}].")
    else:
        # Try case-insensitive match
        matches = [k for k in db["categories"] if k.lower() == category.lower()]
        if matches:
            del db["categories"][matches[0]]
            save(db)
            print(f"Removed category [{matches[0]}].")
        else:
            print(f"Category [{category}] not found.")


def stats() -> None:
    db = load()
    cats = db["categories"]
    total_examples = sum(len(v["examples"]) for v in cats.values())
    print(f"Categories: {len(cats)}")
    print(f"Total trained examples: {total_examples}")


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
        train(args[1], " ".join(args[2:]))

    elif cmd == "classify":
        if len(args) < 2:
            print("Usage: python classifier.py classify <name>")
            sys.exit(1)
        raw = " ".join(args[1:])
        category, standardized, confidence = classify(raw)
        if category is None:
            print(f'Input:        "{raw}"')
            print(f"Category:     UNKNOWN  (best confidence {confidence:.0%})")
        else:
            print(f'Input:        "{raw}"')
            print(f"Standardized: {standardized}")
            print(f"Confidence:   {confidence:.0%}")

    elif cmd == "batch":
        if len(args) < 2:
            print("Usage: python classifier.py batch <file.txt>")
            sys.exit(1)
        path = Path(args[1])
        if not path.exists():
            print(f"File not found: {path}")
            sys.exit(1)
        names = [ln.strip() for ln in path.read_text().splitlines() if ln.strip()]
        results = batch_classify(names)
        print(f"{'PO Name':<60} {'Standardized':<45} {'Conf':>6}")
        print("-" * 115)
        for raw, cat, std, conf in results:
            std_out = std if cat else "** UNKNOWN **"
            print(f"{raw:<60} {std_out:<45} {conf:>5.0%}")

    elif cmd == "list":
        list_categories()

    elif cmd == "reset":
        if len(args) < 2:
            print("Usage: python classifier.py reset <category>")
            sys.exit(1)
        reset_category(" ".join(args[1:]))

    elif cmd == "stats":
        stats()

    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
