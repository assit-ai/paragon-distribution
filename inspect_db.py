import sqlite3

conn = sqlite3.connect('distribution.db')
c = conn.cursor()
tables = [row[0] for row in c.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
for t in sorted(tables):
    try:
        count = c.execute(f'SELECT count(*) FROM "{t}"').fetchone()[0]
        print(f'{t}: {count}')
    except Exception as e:
        print(f'{t}: ERROR {e}')
conn.close()
