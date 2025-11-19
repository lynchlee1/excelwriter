import requests
import zipfile
from io import BytesIO
import re
import json
from config import API_KEY

def get_dart_report(rcept_no: str):
    url = f"https://opendart.fss.or.kr/api/document.xml?crtfc_key={API_KEY}&rcept_no={rcept_no}"
    response = requests.get(url)
    return response.content

def unpack_zip(data):
    with zipfile.ZipFile(BytesIO(data)) as zf:
        content = zf.read(zf.namelist()[0])
    try:
        return content.decode('utf-8')
    except UnicodeDecodeError:
        return content.decode('cp949', errors='ignore')

TABLE_PATTERN = re.compile(r'<TABLE[^>]*>.*?</TABLE>', re.DOTALL | re.IGNORECASE)
ROW_PATTERN = re.compile(r'<TR[^>]*>(.*?)</TR>', re.DOTALL | re.IGNORECASE)
CELL_PATTERN = re.compile(r'<T[DH][^>]*>(.*?)</T[DH]>', re.DOTALL | re.IGNORECASE)

def _clean_html(text):
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', text)).strip()

def _split_paragraphs(text):
    blocks = []
    for para in [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]:
        cleaned = _clean_html(para)
        parts = [chunk.strip() for chunk in re.split(r'(?<=[.?!])\s+', cleaned) if chunk.strip()]
        if parts:
            blocks.append(parts)
    return blocks

def _parse_table(html):
    rows = []
    for row in ROW_PATTERN.findall(html):
        cells = [_clean_html(cell) for cell in CELL_PATTERN.findall(row)]
        if cells:
            rows.append(cells)
    return rows

def split_texts(text):
    paragraphs = []
    tables = []
    last = 0
    for match in TABLE_PATTERN.finditer(text):
        before = text[last:match.start()]
        if before.strip():
            paragraphs.extend(_split_paragraphs(before))
        table_rows = _parse_table(match.group(0))
        if table_rows:
            tables.append(table_rows)
        last = match.end()
    tail = text[last:]
    if tail.strip():
        paragraphs.extend(_split_paragraphs(tail))
    return {"paragraphs": paragraphs, "tables": tables}

def extract_sections(text):
    matches = list(re.finditer(r'<TITLE[^>]*>(.*?)</TITLE>', text, re.DOTALL | re.IGNORECASE))
    if not matches: return None
    sections = {}
    for i, match in enumerate(matches):
        title = re.sub(r'<[^>]+>', '', match.group(1)).strip()
        if title:
            end = matches[i+1].start() if i+1 < len(matches) else len(text)
            sections[title.replace(' ', '')] = split_texts(text[match.end():end].strip())
    return sections

def search_text(node, keyword):
    matches = []
    if isinstance(node, dict):
        for key, value in node.items():
            if isinstance(key, str) and keyword in key:
                matches.append(value)
            matches.extend(search_text(value, keyword))
    elif isinstance(node, list):
        for item in node:
            matches.extend(search_text(item, keyword))
    elif isinstance(node, str):
        if keyword in node:
            matches.append(node)
    return matches

# with open("20251103000190.json", 'w', encoding='utf-8') as f:
#     json.dump(
#         extract_sections(unpack_zip(get_dart_report("20251103000190"))),
#         f,
#         ensure_ascii=False,
#         indent=2
#     )

with open("20251103000190.json", 'r', encoding='utf-8') as f:
    data = json.load(f)

# results = search_text(data, "기타위험")
# results = search_text(results, "희석")
# results = search_text(data, "공모개요")
# results = search_text(results, "희석")

results = search_text(data, "공모방법")
results = search_text(data, "공모방법")


for result in results:
    print(result)
    print("-"*100)
