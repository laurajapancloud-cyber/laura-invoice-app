
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
for i in range(1700-1, 2610):
    line = content[i]
    o = line.count('{')
    c = line.count('}')
    level += o - c
    if o > 0 or c > 0:
        print(f"{i+1:4d}: {level:2d} | {line.strip()}")
