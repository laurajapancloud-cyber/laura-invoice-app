
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
for i in range(1710, 3600):
    line = content[i]
    if line.strip() and not line.startswith('      '):
        # Found a line with less than 6 spaces of indentation
        # (Assuming 6 spaces is the nested level)
        if '}' in line or '@media' in line:
            print(f"{i+1:4d}: {line}")
