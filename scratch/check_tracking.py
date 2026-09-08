with open('templates/index.html', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

for i, l in enumerate(lines, 1):
    if 'data-tab' in l or 'nav-tab' in l or 'switchTab' in l:
        if i > 2900 and i < 3100:
            print(f"{i}: {l.strip()[:140]}")
