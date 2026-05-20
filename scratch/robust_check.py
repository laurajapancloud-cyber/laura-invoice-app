
import sys
content = open('templates/index.html', 'r', encoding='utf-8').read()
import re

# Find the style block
style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
if not style_match:
    print("No style block found")
    sys.exit(0)

style_css = style_match.group(1)
style_start_pos = style_match.start(1)

# Function to get line number from character position
def get_line_no(pos):
    return content.count('\n', 0, pos) + 1

stack = []
for match in re.finditer(r'\{|\}', style_css):
    char = match.group()
    pos = match.start() + style_start_pos
    line_no = get_line_no(pos)
    
    if char == '{':
        stack.append(line_no)
    else:
        if not stack:
            print(f"Extra '}}' at line {line_no}")
        else:
            stack.pop()

if stack:
    for s in stack:
        print(f"Unclosed '{{' at line {s}")
