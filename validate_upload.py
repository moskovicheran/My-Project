# -*- coding: utf-8 -*-
"""Pre-flight check for a PPPoker export BEFORE it is uploaded to the app.

The importer (app/routes/upload.py) reads this file by FIXED sheet names,
a FIXED first data row and FIXED column numbers. A report whose header
changed can therefore still import "successfully" and be silently wrong.
This script checks every assumption the importer makes, and compares the
totals against a reference file (yesterday's export) so a missing block
shows up as a number, not as a surprise next week.

    python validate_upload.py NEW.xlsx [REFERENCE.xlsx]

Exit code 0 = safe to upload, 1 = do not upload.
"""
import sys
import pandas as pd

MEMBER = 'Union Member Statistics'
OVERVIEW = 'Union Overview'
MTT_STATS = 'Union MTT Statistics'
NEEDED_SOFT = ['Union Ring Game Detail', 'Union MTT Detail', MTT_STATS]

FIRST_DATA_ROW = 6          # upload.py: `for i in range(6, len(df))`
COL = {2: 'Super Agent ID', 3: 'Super Agent Nickname', 4: 'Agent ID',
       5: 'Agent Nickname', 7: 'Role', 8: 'Member ID', 9: 'Member Nickname',
       37: 'Player P&L Total', 64: 'Rake&Fee Total'}
MTT_RAKE_COLS = range(56, 64)   # Rake&Fee / MTT / NLH..6+
RING_RAKE_COLS = range(38, 48)  # Ring Game + AoF
SNG_RAKE_COLS = range(48, 56)

fails, warns = [], []


def ok(msg):
    print('  [OK]   ' + msg)


def warn(msg):
    print('  [WARN] ' + msg)
    warns.append(msg)


def fail(msg):
    print('  [FAIL] ' + msg)
    fails.append(msg)


def num(df, cols, rows=None):
    f = df if rows is None else df[rows]
    tot = 0.0
    for c in cols:
        if c < f.shape[1]:
            tot += pd.to_numeric(f[c], errors='coerce').fillna(0).sum()
    return float(tot)


def load(path):
    return pd.read_excel(path, sheet_name=None, header=None)


def member_rows(df):
    """The same rows the importer would store: a member id and a nickname,
    and not the per-club TOTAL line."""
    return (df[8].notna() & df[9].notna()
            & df[1].astype(str).str.upper().ne('TOTAL')
            & df.index.to_series().ge(FIRST_DATA_ROW))


def check_file(path):
    print('\n=== %s ===' % path)
    sheets = load(path)

    # ── A. sheets ───────────────────────────────────────────────────────
    print('\nA. גיליונות')
    if MEMBER not in sheets:
        fail('חסר הגיליון "%s" — הייבוא יחזיר 0 ולא ייכנס כלום' % MEMBER)
        return None, sheets
    ok('"%s" קיים' % MEMBER)

    if OVERVIEW not in sheets:
        fail('חסר "%s" — התאריך ייקבע לפי היום, לא לפי הדוח' % OVERVIEW)
    else:
        raw = str(sheets[OVERVIEW].iloc[2, 0])
        part = raw.replace('Period : ', '').strip().split(' ')[0].split('~')[0].strip()
        try:
            d = pd.to_datetime(part, format='%Y-%m-%d').date()
            ok('תאריך הדוח נקרא: %s  (מתוך "%s")' % (d, raw[:46]))
        except Exception:
            fail('שורה 3 של "%s" לא נקראת כתאריך: %r — הייבוא ישתיל את '
                 'תאריך היום בשקט' % (OVERVIEW, raw[:60]))

    for s in NEEDED_SOFT:
        (ok if s in sheets else warn)(
            '"%s" %s' % (s, 'קיים' if s in sheets else 'חסר'))

    df = sheets[MEMBER]

    # ── B. layout ───────────────────────────────────────────────────────
    print('\nB. מבנה הגיליון')
    print('       גודל: %d שורות x %d עמודות' % df.shape)
    hdr = df.iloc[:FIRST_DATA_ROW]
    labels = {}
    for c in range(df.shape[1]):
        parts = [str(hdr.iloc[r, c]) for r in range(3, FIRST_DATA_ROW)
                 if r < len(hdr)]
        labels[c] = ' / '.join(p for p in parts if p != 'nan')

    for c, expect in COL.items():
        got = labels.get(c, '')
        tail = expect.split()[-1].lower()
        if not got:
            fail('עמודה %d ריקה בכותרת — צפוי "%s"' % (c, expect))
        elif tail in got.lower() or expect.lower() in got.lower():
            ok('עמודה %-3d = %s' % (c, got))
        else:
            fail('עמודה %d היא "%s" — הייבוא מצפה ל-"%s"' % (c, got, expect))

    last = df.shape[1] - 1
    if 'total' in labels.get(last, '').lower():
        ok('עמודה אחרונה (%d) = %s  ← ממנה נקראות הידיים' % (last, labels[last]))
    else:
        fail('העמודה האחרונה (%d) היא "%s" ולא Total — מספר הידיים ייקרא '
             'מהעמודה הלא נכונה' % (last, labels.get(last, '')))

    mtt_labels = [labels.get(c, '') for c in MTT_RAKE_COLS]
    if all(x for x in mtt_labels):
        ok('עמודות רייק הטורנירים (56–63): %s' % ', '.join(
            x.split('/')[-1].strip() for x in mtt_labels))
    else:
        fail('עמודות 56–63 (רייק טורנירים) לא מזוהות — הפיצול ייצא 0')

    clubs = df[0].astype(str).str.contains(r'\(ID:', regex=True, na=False)
    if clubs.sum():
        ok('%d שורות מועדון בפורמט "שם (ID:מספר)"' % clubs.sum())
    else:
        fail('אין שורות מועדון בפורמט "(ID:" — כל השחקנים ייכנסו בלי מועדון')

    rows = member_rows(df)
    if rows.sum():
        ok('%d שורות שחקן יוכנסו' % rows.sum())
    else:
        fail('לא זוהתה אף שורת שחקן')

    # ── C. the arithmetic the split depends on ──────────────────────────
    print('\nC. בדיקות סכום')
    rake = num(df, [64], rows)
    mtt = num(df, MTT_RAKE_COLS, rows)
    parts_sum = num(df, list(RING_RAKE_COLS) + list(SNG_RAKE_COLS) +
                    list(MTT_RAKE_COLS), rows)
    if rake and abs(parts_sum - rake) < max(1.0, rake * 0.001):
        ok('פירוק הרייק מסתכם לעמודת Total: %.2f מול %.2f' % (parts_sum, rake))
    else:
        fail('פירוק הרייק (%.2f) לא שווה לעמודת Total (%.2f) — עמודה 64 '
             'כבר לא הסכום, והפיצול יהיה שגוי' % (parts_sum, rake))

    if MTT_STATS in sheets:
        t = sheets[MTT_STATS]
        tr = t[8].notna() & t[9].notna() if t.shape[1] > 9 else None
        fee_entries = 0.0
        if t.shape[1] > 17:
            body = t.iloc[FIRST_DATA_ROW - 1:]
            fee_entries = float(
                (pd.to_numeric(body[9], errors='coerce').fillna(0) *
                 pd.to_numeric(body[17], errors='coerce').fillna(0)).sum())
        if abs(fee_entries - mtt) < max(1.0, max(mtt, 1) * 0.02):
            ok('רייק הטורנירים תואם את גיליון הטורנירים: %.2f מול %.2f '
               '(עמלה x כניסות)' % (mtt, fee_entries))
        else:
            warn('רייק הטורנירים %.2f לא תואם עמלה x כניסות %.2f — ייתכן '
                 'טורניר שעוד פעיל, או ששתי העמודות זזו' % (mtt, fee_entries))

    print('       רייק כולל %.2f  ·  מזה טורנירים %.2f (%.1f%%)'
          % (rake, mtt, 100 * mtt / rake if rake else 0))
    return df, sheets


def compare(new_df, ref_df):
    print('\nD. השוואה לקובץ הייחוס')
    nr, rr = member_rows(new_df), member_rows(ref_df)
    metrics = [
        ('שורות שחקן', float(nr.sum()), float(rr.sum())),
        ('מועדונים', float(new_df[0].astype(str).str.contains(r'\(ID:', regex=True, na=False).sum()),
         float(ref_df[0].astype(str).str.contains(r'\(ID:', regex=True, na=False).sum())),
        ('עמודות', float(new_df.shape[1]), float(ref_df.shape[1])),
        ('רייק כולל', num(new_df, [64], nr), num(ref_df, [64], rr)),
        ('רייק טורנירים', num(new_df, MTT_RAKE_COLS, nr),
         num(ref_df, MTT_RAKE_COLS, rr)),
        ('P&L כולל', num(new_df, [37], nr), num(ref_df, [37], rr)),
        ('ידיים', num(new_df, [new_df.shape[1] - 1], nr),
         num(ref_df, [ref_df.shape[1] - 1], rr)),
    ]
    print('       %-16s %14s %14s %10s' % ('', 'חדש', 'ייחוס', 'שינוי'))
    for name, a, b in metrics:
        pct = (100 * (a - b) / b) if b else (0 if a == 0 else 100)
        print('       %-16s %14.2f %14.2f %9.1f%%' % (name, a, b, pct))
        if name == 'עמודות' and a != b:
            fail('מספר העמודות השתנה (%d -> %d) — כל המיקומים הקבועים זזו'
                 % (b, a))
        elif b and a == 0:
            fail('%s התאפס בקובץ החדש' % name)
        elif b and abs(pct) > 80 and name not in ('P&L כולל',):
            warn('%s זז ב-%.0f%% — לבדוק שזה אמיתי ולא בלוק חסר' % (name, pct))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    new_df, _ = check_file(sys.argv[1])
    if new_df is not None and len(sys.argv) > 2:
        ref_df, _ = check_file(sys.argv[2])
        if ref_df is not None:
            compare(new_df, ref_df)

    print('\n' + '=' * 62)
    if fails:
        print('אל תעלה את הקובץ — %d בעיות חוסמות:' % len(fails))
        for f in fails:
            print('  - ' + f)
    if warns:
        print('%d אזהרות (לא חוסמות):' % len(warns))
        for w in warns:
            print('  - ' + w)
    if not fails:
        print('הקובץ תואם למה שהייבוא מצפה לו. אפשר להעלות.')
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main())
