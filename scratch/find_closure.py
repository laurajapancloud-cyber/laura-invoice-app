
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
for i in range(1710-1, 2467):
    line = content[i]
    o = line.count('{')
    c = line.count('}')
    level += o - c
    if level == 0 and (o > 0 or c > 0):
        print(f"Media query closed at line {i+1}: {line.strip()}")
print("Done")
