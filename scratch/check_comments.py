
import re
content = open('templates/index.html', 'r', encoding='utf-8').read()
styles = re.findall(r'<style>(.*?)</style>', content, re.DOTALL)
for i, s in enumerate(styles):
    opens = s.count('/*')
    closes = s.count('*/')
    if opens != closes:
        print(f"Style block {i} comment mismatch: {opens} opens vs {closes} closes")
        # Find which one is unclosed
        # (This is a bit simple but usually works)
        pass
print("Done")
