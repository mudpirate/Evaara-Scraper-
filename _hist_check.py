import datetime
import os
import shutil
import sqlite3

EPOCH = datetime.datetime(1601, 1, 1, tzinfo=datetime.timezone.utc)

prof = os.path.expanduser(r"~\.cache\maps-scraper\profile\Default\History")
tmp = os.path.join(os.environ["TEMP"], "maps_history_copy.db")
shutil.copy(prof, tmp)

con = sqlite3.connect(f"file:{tmp}?mode=ro", uri=True)

print("=== urls (by last_visit_time, top 20 with time) ===")
for ti, u in con.execute(
    "select last_visit_time, url from urls order by last_visit_time desc limit 20"
).fetchall():
    dt = (EPOCH + datetime.timedelta(microseconds=ti)).astimezone()
    print(dt.strftime("%Y-%m-%d %H:%M:%S"), u[:100])

print()
print("=== visits (last 30) ===")
try:
    for vt, u in con.execute(
        "select v.visit_time, u.url from visits v join urls u on v.url = u.id "
        "order by v.visit_time desc limit 30"
    ).fetchall():
        dt = (EPOCH + datetime.timedelta(microseconds=vt)).astimezone()
        print(dt.strftime("%Y-%m-%d %H:%M:%S"), u[:100])
except Exception as e:
    print("visits err:", e)

con.close()
os.remove(tmp)