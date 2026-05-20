
import sys
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
start = 3000
end = 3600
for i in range(start-1, end):
    line = content[i]
    old_level = level
    level += line.count('{') - line.count('}')
    if level != old_level or '{' in line or '}' in line:
        print(f'{i+1:4d}: {level:2d} | {line.strip()}')
