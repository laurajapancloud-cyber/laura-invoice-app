
import re
content = open('templates/index.html', 'r', encoding='utf-8').read()
styles = re.findall(r'<style>(.*?)</style>', content, re.DOTALL)
for i, s in enumerate(styles):
    # This is a bit naive because of escaped quotes, but usually works
    for quote in ["'", '"']:
        if s.count(quote) % 2 != 0:
            print(f"Style block {i} unclosed quote {quote}")
print("Done")
