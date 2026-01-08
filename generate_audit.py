#!/usr/bin/env python3

import argparse
import os
import re
from typing import Dict, List, Optional, Set, Tuple

from unidecode import unidecode


PLACE_RE = re.compile(r"^(\d+)\.(.*)$")


def _norm(name: str) -> str:
    return " ".join(unidecode(name).lower().split())


def read_names_from_results(path: str) -> List[str]:
    names: List[str] = []
    if not os.path.exists(path):
        return names

    with open(path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line:
                continue
            m = PLACE_RE.match(line)
            rest = m.group(2).strip() if m else line
            for part in rest.split("|"):
                p = part.strip()
                if p:
                    names.append(p)
    return names


def write_audit(
    root: str,
    out_path: str,
    only_dirs: Optional[Set[str]] = None,
    only_tournaments: Optional[Set[str]] = None,
) -> None:
    blocks: List[str] = []

    # If an explicit tournament allowlist is provided, use that.
    if only_tournaments is not None:
        targets: List[Tuple[str, str]] = []
        for rel in sorted(only_tournaments):
            rel = rel.strip().strip("/")
            if not rel:
                continue
            parts = rel.split("/", 1)
            if len(parts) != 2:
                continue
            targets.append((parts[0], parts[1]))
    else:
        # Otherwise scan for tournament directories containing mens.txt
        targets = []
        for club in sorted(
            d
            for d in os.listdir(root)
            if os.path.isdir(os.path.join(root, d)) and not d.startswith(".")
        ):
            if only_dirs is not None and club not in only_dirs:
                continue
            club_path = os.path.join(root, club)
            for tname in sorted(
                d for d in os.listdir(club_path) if os.path.isdir(os.path.join(club_path, d))
            ):
                targets.append((club, tname))

    for club, tname in targets:
        club_path = os.path.join(root, club)
        tdir = os.path.join(club_path, tname)
        mens_path = os.path.join(tdir, "mens.txt")
        if not os.path.exists(mens_path):
            continue

        womens_path = os.path.join(tdir, "womens.txt")
        juniors_path = os.path.join(tdir, "juniors.txt")
        junior_path = os.path.join(tdir, "junior.txt")
        mixed_path = os.path.join(tdir, "mixed.txt")

        men = read_names_from_results(mens_path)
        women = read_names_from_results(womens_path)
        juniors = read_names_from_results(juniors_path)
        if not juniors:
            juniors = read_names_from_results(junior_path)
        mixed = read_names_from_results(mixed_path)

        # unique, preserve order
        def uniq(seq: List[str]) -> List[str]:
            seen: Set[str] = set()
            out: List[str] = []
            for x in seq:
                k = _norm(x)
                if k in seen:
                    continue
                seen.add(k)
                out.append(x)
            return out

        men_u = uniq(men)
        women_u = uniq(women)
        juniors_u = uniq(juniors)
        mixed_u = uniq(mixed)

        blocks.append(f"=== {club}/{tname} ===")
        blocks.append(f"men({len(men_u)}): " + (", ".join(men_u) if men_u else ""))
        blocks.append(f"women({len(women_u)}): " + (", ".join(women_u) if women_u else ""))
        blocks.append(f"junior({len(juniors_u)}): " + (", ".join(juniors_u) if juniors_u else ""))
        blocks.append(f"mixed({len(mixed_u)}): " + (", ".join(mixed_u) if mixed_u else ""))
        blocks.append("")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("TOURNAMENT ROSTERS (men/women/junior/mixed):\n\n")
        f.write("\n".join(blocks).rstrip() + "\n")


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate audit file listing rosters per tournament from .txt outputs")
    ap.add_argument("--root", default=".")
    ap.add_argument("--out", default="audit_2026-01-08_all_clubs.txt")
    ap.add_argument(
        "--clubs",
        default="",
        help="Optional comma-separated list of club directories to include (e.g. 'wolfpack,spacerebels')",
    )
    ap.add_argument(
        "--only-tournaments",
        default="",
        help=(
            "Optional comma-separated list of explicit tournament paths to include, in the form "
            "'club/tournament-folder'. If set, this overrides --clubs scanning."
        ),
    )
    args = ap.parse_args()

    only_dirs: Optional[Set[str]] = None
    if args.clubs.strip():
        only_dirs = {c.strip() for c in args.clubs.split(",") if c.strip()}

    only_tournaments: Optional[Set[str]] = None
    if args.only_tournaments.strip():
        only_tournaments = {t.strip() for t in args.only_tournaments.split(",") if t.strip()}

    write_audit(
        args.root,
        os.path.join(args.root, args.out),
        only_dirs=only_dirs,
        only_tournaments=only_tournaments,
    )


if __name__ == "__main__":
    main()
