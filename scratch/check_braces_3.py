
import sys
content = open('templates/index.html', 'r', encoding='utf-8').read().splitlines()
level = 0
start = 2990
end = 3010
for i in range(start-1, end):
    line = content[i]
    old_level = level
    level += line.count('{') - line.count('}'); print(f'{i+1:4d}: {level:2d} | {line.strip()}')
