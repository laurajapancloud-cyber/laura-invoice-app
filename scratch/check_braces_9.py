
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
for i in range(1710-1, 3000):
    line = content[i]
    o = line.count('{')
    c = line.count('}')
    old_level = level
    level += o - c
    if level != old_level or o > 0 or c > 0:
        print(f"{i+1:4d}: {level:2d} | {line.strip()}")
