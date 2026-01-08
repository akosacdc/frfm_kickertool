#!/usr/bin/env python3

import argparse
import os
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

from unidecode import unidecode


PLACE_RE = re.compile(r"^(\d+)\.(.*)$")


def _normalize_for_match(name: str) -> str:
    # Match should be robust to accents/case/multiple spaces.
    return " ".join(unidecode(name).lower().split())


def _swap_first_last_if_two_tokens(norm_name: str) -> Optional[str]:
    parts = norm_name.split()
    if len(parts) != 2:
        return None
    return f"{parts[1]} {parts[0]}"


def _pretty_token(token: str) -> str:
    # Handle hyphenated tokens; preserve non-letter characters.
    parts = token.split("-")
    pretty_parts = []
    for p in parts:
        if not p:
            pretty_parts.append(p)
            continue
        pretty_parts.append(p[:1].upper() + p[1:].lower())
    return "-".join(pretty_parts)


def _entry_key(players: List[str]) -> str:
    # Key for singles (len=1) or doubles (len=2) entries, preserving order.
    return "|".join(_normalize_for_match(p) for p in players)


def pretty_name(name: str) -> str:
    name = " ".join(name.split())
    # If the whole name is ALL CAPS (after removing non-letters), title-case it.
    letters_only = re.sub(r"[^A-Za-z]", "", unidecode(name))
    if letters_only and letters_only.isupper():
        return " ".join(_pretty_token(tok) for tok in name.split(" "))

    # Otherwise, still normalize weird mixed casing like "ALEX tOne".
    return " ".join(_pretty_token(tok) for tok in name.split(" "))


@dataclass(frozen=True)
class Placement:
    place: int
    players: List[str]


def parse_results_txt(lines: Iterable[str]) -> List[Placement]:
    placements: List[Placement] = []
    last_place: Optional[int] = None

    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        m = PLACE_RE.match(line)
        if m:
            place = int(m.group(1))
            rest = m.group(2).strip()
            last_place = place
        else:
            if last_place is None:
                # Malformed file; just skip.
                continue
            place = last_place
            rest = line

        # Doubles in this project are encoded as "Name1|Name2".
        players = [p.strip() for p in rest.split("|") if p.strip()]
        placements.append(Placement(place=place, players=players))

    return placements


def _ensure_explicit_places_for_simple_list(lines: List[str]) -> List[str]:
    """
    If a file effectively contains a simple ordered list but lacks explicit place numbers
    (common bug: only the first line is like "1.Name", rest are bare), assign 1..n.
    This intentionally does NOT try to infer ties.
    """
    cleaned = [ln.strip() for ln in lines if ln.strip()]
    if not cleaned:
        return []

    numbered_count = sum(1 for ln in cleaned if PLACE_RE.match(ln))
    if numbered_count == 1 and PLACE_RE.match(cleaned[0]):
        # This pattern is ambiguous: it may mean everyone tied on place 1.
        # Do NOT invent sequential ranks; just normalize casing.
        out: List[str] = []
        m0 = PLACE_RE.match(cleaned[0])
        assert m0 is not None
        out.append(f"1.{pretty_name(m0.group(2).strip())}")
        for ln in cleaned[1:]:
            m = PLACE_RE.match(ln)
            if m:
                out.append(f"{int(m.group(1))}.{pretty_name(m.group(2).strip())}")
            else:
                out.append(pretty_name(ln))
        return out

    return cleaned


def render_results_txt(placements: List[Placement]) -> List[str]:
    out: List[str] = []
    prev_place: Optional[int] = None

    for pl in placements:
        players_pretty = "|".join(pretty_name(p) for p in pl.players)
        if prev_place is None or pl.place != prev_place:
            out.append(f"{pl.place}.{players_pretty}")
        else:
            out.append(players_pretty)
        prev_place = pl.place

    return out


def build_open_place_index(open_placements: List[Placement]) -> Dict[str, int]:
    idx: Dict[str, int] = {}
    for pl in open_placements:
        # Index the full entry (single or doubles team)
        entry_key = _entry_key(pl.players)
        if entry_key and (entry_key not in idx or pl.place < idx[entry_key]):
            idx[entry_key] = pl.place
        if len(pl.players) == 2:
            swapped_team = _entry_key([pl.players[1], pl.players[0]])
            if swapped_team and (swapped_team not in idx or pl.place < idx[swapped_team]):
                idx[swapped_team] = pl.place

        for p in pl.players:
            key = _normalize_for_match(p)
            # If a name appears multiple times, keep the best (lowest) place.
            if key and (key not in idx or pl.place < idx[key]):
                idx[key] = pl.place

            swapped = _swap_first_last_if_two_tokens(key)
            if swapped and (swapped not in idx or pl.place < idx[swapped]):
                idx[swapped] = pl.place
    return idx


def rerank_subcategory(
    sub_lines: List[str],
    open_idx: Dict[str, int],
    filename_for_logs: str,
) -> List[str]:
    # Parse to preserve any explicit numbering/ties.
    sub_placements = parse_results_txt(sub_lines)
    raw_entries: List[List[str]] = [pl.players for pl in sub_placements]

    # Map each player to their open-place.
    entries: List[Tuple[int, List[str]]] = []
    missing: List[str] = []
    for players in raw_entries:
        open_place: Optional[int] = None

        # Prefer matching by full entry (singles or doubles team)
        ek = _entry_key(players)
        if ek:
            open_place = open_idx.get(ek)

        # For singles, also try swapped first/last
        if open_place is None and len(players) == 1:
            key = _normalize_for_match(players[0])
            swapped = _swap_first_last_if_two_tokens(key)
            if swapped is not None:
                open_place = open_idx.get(swapped)

        # For doubles, also try swapped team order
        if open_place is None and len(players) == 2:
            swapped_team = _entry_key([players[1], players[0]])
            if swapped_team:
                open_place = open_idx.get(swapped_team)

        if open_place is None:
            missing.append("|".join(players))
            # Put missing players last in stable order.
            open_place = 10**9
        entries.append((open_place, players))

    # If nothing matches the open standings, this is likely a separate women/junior event.
    # In that case, keep the existing placements and just normalize output formatting.
    if len(missing) == len(raw_entries):
        rendered = render_results_txt(sub_placements)
        # If the file looked like a simple list with missing numbers, fix that too.
        return _ensure_explicit_places_for_simple_list(rendered)

    if missing:
        print(f"[WARN] {filename_for_logs}: missing in open results: {missing}")

    # Sort by open place, stable.
    entries_sorted = sorted(enumerate(entries), key=lambda t: (t[1][0], t[0]))
    sorted_entries = [entries[i][1] for i, _ in entries_sorted]
    sorted_open_places = [entries[i][0] for i, _ in entries_sorted]

    # Convert open-place ties into subcategory ranks 1..n with ties.
    out: List[str] = []
    rank = 0
    last_open_place: Optional[int] = None

    for open_place, players in zip(sorted_open_places, sorted_entries):
        if last_open_place is None or open_place != last_open_place:
            rank += 1
            out.append(f"{rank}.{'|'.join(pretty_name(p) for p in players)}")
        else:
            out.append("|".join(pretty_name(p) for p in players))
        last_open_place = open_place

    return out


def process_tournament_dir(tournament_dir: str, dry_run: bool) -> bool:
    changed_any = False

    mens_path = os.path.join(tournament_dir, "mens.txt")
    if not os.path.exists(mens_path):
        return False

    with open(mens_path, "r", encoding="utf-8") as f:
        mens_lines = f.read().splitlines()

    mens_placements = parse_results_txt(mens_lines)
    open_idx = build_open_place_index(mens_placements)

    # Always normalize name casing in mens.txt too.
    mens_new_lines = render_results_txt(mens_placements)
    if mens_new_lines != [l.strip() for l in mens_lines if l.strip()]:
        changed_any = True
        if not dry_run:
            with open(mens_path, "w", encoding="utf-8") as f:
                f.write("\n".join(mens_new_lines) + "\n")

    for sub in ("womens.txt", "juniors.txt", "junior.txt"):
        sub_path = os.path.join(tournament_dir, sub)
        if not os.path.exists(sub_path):
            continue

        with open(sub_path, "r", encoding="utf-8") as f:
            sub_lines = f.read().splitlines()

        new_lines = rerank_subcategory(sub_lines, open_idx, sub_path)

        # Also normalize formatting even if reranking does not change order.
        if new_lines != [l.strip() for l in sub_lines if l.strip()]:
            changed_any = True
            if not dry_run:
                with open(sub_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(new_lines) + "\n")

    return changed_any


def iter_tournament_dirs(root: str) -> Iterable[str]:
    for dirpath, _, filenames in os.walk(root):
        if "mens.txt" in filenames:
            yield dirpath


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Fix generated tournament .txt files: normalize ALL-CAPS names and "
            "re-rank womens/juniors by their open placement from mens.txt (ties preserved)."
        )
    )
    ap.add_argument("--root", default=".", help="Root folder to scan (default: current dir)")
    ap.add_argument("--dry-run", action="store_true", help="Do not write, just report")
    args = ap.parse_args()

    changed_dirs: List[str] = []
    for tdir in iter_tournament_dirs(args.root):
        if process_tournament_dir(tdir, args.dry_run):
            changed_dirs.append(tdir)

    print(f"Processed {len(changed_dirs)} tournament directories with changes")


if __name__ == "__main__":
    main()
