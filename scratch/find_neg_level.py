
import re
content = open('templates/index.html', 'r', encoding='utf-8').read()
style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
style_css = style_match.group(1)
level = 0
for i, char in enumerate(style_css):
    if char == '{':
        level += 1
    elif char == '}':
        level -= 1
        if level < 0:
            # Found it!
            # Get line number
            before = style_css[:i]
            line_no = before.count('\n') + style_match.start(0) # approximate
            # Actually better:
            full_before = content[:style_match.start(1) + i]
            actual_line = full_before.count('\n') + 1
            print(f"Level went to -1 at line {actual_line}: {style_css[i-10:i+10].strip()}")
            level = 0
