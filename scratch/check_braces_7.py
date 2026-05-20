
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
for i in range(0, 1710):
    line = content[i]
    o = line.count('{')
    c = line.count('}')
    level += o - c
    if level < 0:
        print(f"{i+1:4d}: {level:2d} | {line.strip()}")
        level = 0 # reset to find more
