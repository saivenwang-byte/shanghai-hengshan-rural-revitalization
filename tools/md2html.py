import re, os

md_path = "docs/横山村乡村振兴项目运营统筹方案.md"
with open(md_path, "r", encoding="utf-8") as f:
    md = f.read()

def md2html(text):
    text = re.sub(r"^#### (.+)", r"<h4>\1</h4>", text, flags=re.M)
    text = re.sub(r"^### (.+)", r"<h3>\1</h3>", text, flags=re.M)
    text = re.sub(r"^## (.+)", r"<h2>\1</h2>", text, flags=re.M)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    lines = text.split("\n")
    result, in_table, after_th = [], False, False
    for line in lines:
        s = line.strip()
        if s.startswith("|") and s.endswith("|"):
            if not in_table:
                result.append('<table>')
                in_table, after_th = True, False
            if "---" in s:
                after_th = True
                continue
            cells = [c.strip() for c in s.split("|")[1:-1]]
            tag = "td" if after_th else "th"
            result.append("<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>")
        else:
            if in_table:
                result.append("</table><br>")
                in_table = False
            if s.startswith("- ") or s.startswith("* "):
                result.append(f"<li>{s[2:]}</li>")
            elif s.startswith("> "):
                result.append(f'<blockquote>{s[2:]}</blockquote>')
            elif s and not s.startswith("#"):
                if "http" in s:
                    s = re.sub(r"(https?://\S+)", r'<a href="\1">\1</a>', s)
                result.append(f"<p>{s}</p>")
            elif not s:
                result.append("<br>")
    if in_table:
        result.append("</table>")
    return "\n".join(result)

html_body = md2html(md)
html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="utf-8">
<style>
body{{font-family:"Microsoft YaHei","SimSun",sans-serif;font-size:11pt;line-height:1.8;max-width:820px;margin:30px auto;color:#333;padding:0 20px}}
h1{{font-size:22pt;text-align:center;color:#1A1A2E;border-bottom:3px solid #E6A817;padding-bottom:10px}}
h2{{font-size:16pt;color:#2D6A4F;margin-top:28px;border-bottom:1px solid #ddd;padding-bottom:4px;page-break-before:always}}
h2:first-of-type{{page-break-before:auto}}
h3{{font-size:13pt;color:#444;margin-top:20px}}
h4{{font-size:11pt;color:#666}}
table{{width:100%;margin:12px 0;font-size:9pt;border-collapse:collapse}}
th{{background:#2D6A4F;color:white;padding:6px 8px;text-align:left}}
td{{padding:5px 7px;border-bottom:1px solid #eee}}
tr:nth-child(even) td{{background:#f9f9f9}}
blockquote{{border-left:4px solid #E6A817;padding:8px 12px;margin:8px 0;background:#fafaf5;color:#666;font-style:italic}}
li{{margin:2px 0}}
a{{color:#2D6A4F}}
@media print{{body{{font-size:10pt}}h2{{page-break-before:always}}}}
</style></head>
<body>{html_body}</body></html>"""

html_path = "docs/横山村乡村振兴项目运营统筹方案.html"
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html)
print(f"HTML: {html_path} ({len(html)} chars)")
