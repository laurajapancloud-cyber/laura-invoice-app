
import re
content = open('templates/index.html', 'r', encoding='utf-8').read()
style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
if style_match:
    style = style_match.group(1)
    print(f"Opens: {style.count('{')}, Closes: {style.count('}')}")
