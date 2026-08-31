import sqlite3

profiles = {
    "real Default": r"C:\Users\Nomesh\AppData\Local\Google\Chrome\User Data\Default",
    "jd-session-1456": r"C:\Users\Nomesh\.cache\jd-session-1456\Default",
    "jd-session-14752": r"C:\Users\Nomesh\.cache\jd-session-14752\Default",
    "justdial-scraper": r"C:\Users\Nomesh\.cache\justdial-scraper\profile\Default",
}

for label, base in profiles.items():
    ck = base + r"\Network\Cookies"
    try:
        con = sqlite3.connect(f"file:{ck}?mode=ro", uri=True)
        rows = con.execute(
            "select host_key,name,length(value),expires_utc from cookies "
            "where host_key like '%justdial%' order by expires_utc desc"
        ).fetchall()
        con.close()
        print(f"\n=== {label} ===")
        print(f"  path exists: {__import__('os').path.exists(ck)}")
        print(f"  justdial cookies: {len(rows)}")
        for r in rows[:40]:
            print("   ", r)
    except Exception as e:
        print(f"\n=== {label} ===\n  error: {e}")
