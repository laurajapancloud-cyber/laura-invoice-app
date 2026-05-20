
import re
content = open('templates/index.html', 'r', encoding='utf-8').read()
style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
style_css = style_match.group(1)
style_start = style_match.start(1)

def get_line(pos):
    return content.count('\n', 0, pos + style_start) + 1

stack = []
for match in re.finditer(r'@media[^{]*\{|\{|\}', style_css):
    text = match.group()
    pos = match.start()
    
    if text.startswith('@media'):
        stack.append(('media', get_line(pos)))
    elif text == '{':
        stack.append(('brace', get_line(pos)))
    elif text == '}':
        if not stack:
            print(f"Extra '}}' at line {get_line(pos)}")
        else:
            type, line = stack.pop()
            if type == 'media':
                print(f"Media query starting at line {line} ends at line {get_line(pos)}")
print("Done")
