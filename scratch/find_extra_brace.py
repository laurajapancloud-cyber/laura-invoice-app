
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
for i in range(2467-1, 3015):
    line = content[i]
    o = line.count('{')
    c = line.count('}')
    level += o - c
    if level < 0:
        print(f"ERROR: Extra '}}' at line {i+1}: {line.strip()}")
        level = 0
print("Done")
