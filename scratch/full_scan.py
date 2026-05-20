
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
for i, line in enumerate(content):
    o = line.count('{')
    c = line.count('}')
    level += o - c
    if level < 0:
        print(f"ERROR: Level below 0 at line {i+1}: {line.strip()} (Level: {level})")
        # Reset level to 0 to find more errors
        level = 0
    if level > 10:
        print(f"WARNING: High level at line {i+1}: {line.strip()} (Level: {level})")
print("Done")
