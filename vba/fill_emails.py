#!/usr/bin/env python3
"""
Fill the Emails sheet of the .xlsm from the committee's emails.xlsx directory.

Names in the schedule and names in the directory rarely match character for
character ("Mohammad Raed" vs "Mohammad Raed Khirfan", "Saed Bassam Mohamad Ali
Al-Hajjawi" vs "Saed Alhajjawi"), so each name is scored against every
directory entry and only confident matches are written. Everything else is
listed for a human to finish.

Run with --apply to write. Without it, prints the report and changes nothing.
"""

import difflib
import json
import os
import re
import shutil
import sys
import tempfile
import warnings

warnings.filterwarnings("ignore")

import openpyxl
import win32com.client as win32

FOLDER = (
    r"C:\Users\Faris.Alsalem\OneDrive - AL-Hussien bin Abdullah Technical University"
    r"\Faris HTU\Exam Committee"
)
XLSM = os.path.join(FOLDER, "Proctor Schedule by Date.xlsm")
DIRECTORY = os.path.join(FOLDER, "emails.xlsx")

ACCEPT = 75          # below this we do not write anything

# Same person, spelled differently enough that no safe rule catches it. Kept
# explicit rather than loosening the scorer, which would let genuinely
# different people match. Each was checked by eye against the directory.
MANUAL_ALIASES = {
    "Khaled Ahmad Mahmoud Alza'areer": "khalid.alzareer@htu.edu.jo",   # Khalid AlZareer
    "Lena Ali Salem Al-Samameh": "lina.alsamameh@htu.edu.jo",          # Lina AlSamameh
    "Othman Alsmadi": "othman.smadi@htu.edu.jo",                       # Othman Smadi
}
PAIR_RE = re.compile(r'^\s*"?(?P<name>[^"<>]+?)"?\s*<(?P<mail>[^<>@\s]+@[^<>\s]+)>\s*$')

# Titles only. Particles like "al", "abu" and "bani" are part of the surname
# ("Bani Issa" -> baniissa) and dropping them loses the very token that
# identifies the person.
NOISE = {"dr", "prof", "mr", "mrs", "ms", "eng"}


def norm(s):
    s = str(s or "").lower()
    s = s.replace("'", "").replace("`", "").replace("-", " ").replace(".", " ")
    s = re.sub(r"[^a-z ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def tokens(s):
    return [t for t in norm(s).split() if len(t) > 1]


def strong_tokens(s):
    return [t for t in tokens(s) if t not in NOISE]


def load_directory():
    tmp = os.path.join(tempfile.gettempdir(), "_dir.xlsx")
    shutil.copy2(DIRECTORY, tmp)
    wb = openpyxl.load_workbook(tmp, data_only=True)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str):
                    m = PAIR_RE.match(v.strip())
                    if m:
                        mail = m.group("mail").strip().lower()
                        out.setdefault(mail, m.group("name").strip())
    os.remove(tmp)
    return [(name, mail) for mail, name in out.items()]


def keys(s):
    """Tokens plus adjacent concatenations.

    Surnames get split differently on each side - the schedule writes
    "Al Attili" and "Bani Issa" where the directory writes "AlAttili" and
    "BaniIssa" - so joining neighbouring tokens lets the two meet.
    """
    ts = strong_tokens(s)
    out = set(ts)
    for i in range(len(ts) - 1):
        out.add(ts[i] + ts[i + 1])
    if len(ts) > 2:
        out.add("".join(ts[-2:]))
    return out


def score(sheet_name, dir_name, mail):
    """0-100 confidence that dir_name/mail is sheet_name."""
    sn, dn = norm(sheet_name), norm(dir_name)
    if not sn or not dn:
        return 0

    st, dt = strong_tokens(sheet_name), strong_tokens(dir_name)
    if not st or not dt:
        return 0

    if sn == dn:
        return 100

    sset, dset = set(st), set(dt)
    skeys, dkeys = keys(sheet_name), keys(dir_name)
    lset = {t for t in re.split(r"[.\-_]", mail.split("@")[0]) if len(t) > 1}

    # One full name contains the other: "Mohammad Raed" / "Mohammad Raed Khirfan"
    if sset <= dset or dset <= sset:
        return 96

    # Both halves of the address appear in the schedule name. This is the
    # workhorse rule and is strict: the surname must genuinely be present,
    # which is what stopped abdelrahman.yousef matching Al Attili.
    if lset and lset <= skeys:
        return 92

    # First names agree AND a surname-level token agrees. The first name is
    # excluded from the overlap, otherwise every "Ahmad" matches every other.
    if st[0] == dt[0] and ((skeys & dkeys) - {st[0]}):
        return 90

    # Spelling drift: Khaled/Khalid, Alza'areer/AlZareer, Rajaee/Rajaie.
    first_ratio = difflib.SequenceMatcher(None, st[0], dt[0]).ratio()
    if first_ratio < 0.7:
        return 0

    a, b = " ".join(sorted(sset)), " ".join(sorted(dset))
    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    if ratio >= 0.90:
        return 85
    if ratio >= 0.80:
        return 78
    if ratio >= 0.70:
        return 65
    return int(ratio * 60)


def best_match(sheet_name, directory):
    ranked = sorted(
        ((score(sheet_name, dn, mail), dn, mail) for dn, mail in directory),
        key=lambda x: -x[0],
    )
    return ranked[0], ranked[1] if len(ranked) > 1 else (0, "", "")


def main():
    apply = "--apply" in sys.argv
    directory = load_directory()
    print("Directory entries: %d\n" % len(directory))

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    try:
        wb = excel.Workbooks.Open(XLSM)
        ws = wb.Worksheets("Emails")
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row

        ws.Cells(3, 6).Value = "Matched from directory"
        ws.Cells(3, 6).Font.Bold = True
        ws.Cells(3, 6).Interior.Color = 0xF6F1E1
        ws.Columns(6).ColumnWidth = 40

        filled, review, missed = [], [], []

        for r in range(4, last + 1):
            name = str(ws.Cells(r, 1).Value or "").strip()
            if not name:
                continue

            if name in MANUAL_ALIASES:
                mail = MANUAL_ALIASES[name]
                if apply:
                    ws.Cells(r, 3).Value = mail
                    ws.Cells(r, 6).Value = "spelling variant - please verify"
                    ws.Cells(r, 6).Font.Color = 0x9A6A00
                review.append((name, "manual alias", mail, 70))
                continue

            (sc, dn, mail), runner = best_match(name, directory)

            if sc >= ACCEPT:
                if apply:
                    ws.Cells(r, 3).Value = mail
                    ws.Cells(r, 6).Value = "%s  (%d%%)" % (dn, sc)
                    ws.Cells(r, 6).Font.Color = 0x9A9A9A
                row = (name, dn, mail, sc)
                (filled if sc >= 88 else review).append(row)
            else:
                missed.append((name, dn, mail, sc))
                if apply:
                    ws.Cells(r, 6).Value = "no confident match - please fill column C"
                    ws.Cells(r, 6).Font.Color = 0xA62A22
                    ws.Cells(r, 3).Interior.Color = 0xFAE5E3

        print("=== MATCHED (%d) ===" % len(filled))
        for n, d, m, s in filled:
            print("  %3d%%  %-36s -> %s" % (s, n[:36], m))

        print("\n=== MATCHED, WORTH CHECKING (%d) ===" % len(review))
        for n, d, m, s in review:
            print("  %3d%%  %-36s -> %-34s  [directory: %s]" % (s, n[:36], m, d))

        print("\n=== NO CONFIDENT MATCH (%d) - fill these by hand ===" % len(missed))
        for n, d, m, s in missed:
            hint = "closest was %s (%s) at %d%%" % (d, m, s) if d else "nothing close"
            print("  %-36s  %s" % (n[:36], hint))

        if apply:
            wb.Save()
            print("\nWritten to %s" % XLSM)
        else:
            print("\n(dry run - nothing written; re-run with --apply)")
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
