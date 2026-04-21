# סקריפטים חשובים לשמירה

כל סקריפט רץ מול פרוד ע"י הגדרת `$env:DATABASE_URL` לפני ההרצה. דוגמה:
```powershell
$env:DATABASE_URL="<neon-url>"
python <script>.py <args>
Remove-Item env:DATABASE_URL
```

## 1. בדיקת תקינות (Read-only)

### `diag_breakdown.py`
מציג את הדלתא בין Top Box לסכום הכרטיסים. "DELTA = 0.00" → הכל מאוזן.
```powershell
python diag_breakdown.py
```

### `diag_delta_sources.py`
מפרק דלתא קיימת: Orphan rows לפי מועדון, תרומת העברות כספים, overrides לסוכנים שלא בקופסה.
```powershell
python diag_delta_sources.py
```

## 2. שיוך שחקנים (`PlayerAssignment`)

### `tools_assign_player.py` (כללי, חדש)
שיוך / העברה / ביטול של override לשחקן.
```powershell
# שייך שחקן ל‑SA:
python tools_assign_player.py --player-id 1443-8481 --sa-id 8040-6815 --note "SPC T under Dolar 10"

# שייך לסוכן ספציפי:
python tools_assign_player.py --player-id 1234-5678 --agent-id 6670-6318

# בטל שיוך:
python tools_assign_player.py --player-id 1443-8481 --delete
```
Idempotent — אם כבר קיים, יעדכן (לא יכפיל).

## 3. רישום מועדונים לסוכנים (`SARakeConfig`)

### `tools_add_managed_club.py` (כללי, חדש)
רישום / ביטול של מועדון כמנוהל ע"י SA. תומך גם ב‑club_id מה‑Excel וגם בשם ליטרלי (MANG0, Marmalades וכו').
```powershell
# שם ליטרלי:
python tools_add_managed_club.py --sa-id 4406-1298 --club "Marmalades"

# club_id מ‑Excel:
python tools_add_managed_club.py --sa-id 4447-3687 --club "630307"

# מחיקת רישום:
python tools_add_managed_club.py --sa-id 4406-1298 --club "MANG0" --delete
```

## 4. ייצוא שחקן לאקסל

### `tools_export_player.py` (כללי, חדש)
ייצוא מלא של כל הנתונים של שחקן: פירוק לפי סוג משחק, בליינדס, דרך מי שיחק, סשנים, העברות, ועוד.
```powershell
python tools_export_player.py --player-id 1443-8481
python tools_export_player.py --player-id 1443-8481 --out custom.xlsx
```
פלט: `player_<id>_export.xlsx` (או `--out` מותאם).

## 5. סקריפטים חד־פעמיים (אפשר למחוק אחרי שרצו)

קבצים היסטוריים שעשו שינוי אחד ספציפי בפרוד. אין צורך לשמור אותם אם אין צורך בהתאמה נוספת:
- `move_areyoufold_to_niroha.py` — הוחלף ע"י `tools_assign_player.py`
- `add_marmalades_to_mangisto.py` / `add_mang0_to_mangisto.py` — הוחלפו ע"י `tools_add_managed_club.py`
- `dedup_mang0_config.py` / `delete_466566_config.py` — מחיקות חד־פעמיות
- `export_areyoufold.py` / `export_mangisto_2000.py` — הוחלפו ע"י `tools_export_player.py`

## 6. קישורים שימושיים בפנים

- `/admin/health` — דף בדיקת תקינות עם delta, יתומים, double-counts + שיוך ישיר מהרשימה
- `/admin/agents` — ניהול SA hierarchy, SARakeConfig, RakeConfig
- `/admin/lost-players` — שיוך ידני של שחקנים בלי sa_id/agent_id
- `/admin/agent-view/<sa_id>` — דשבורד סוכן (כמנהל)
- `/union/player/<pid>?full=1` — פרטי שחקן (ללא scope filter)
