from tarfile import data_filter
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

from webDataParser import unpack_zip, search, search_tables, clean_text
from webDataParser import TABLE_PATTERN, ROW_PATTERN, CELL_PATTERN

def _split_paragraphs(text):
    blocks = []
    for para in [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]:
        cleaned = clean_text(para)
        parts = [chunk.strip() for chunk in re.split(r'(?<=[.?!])\s+', cleaned) if chunk.strip()]
        if parts:
            blocks.append(parts)
    return blocks

def _parse_table(html):
    rows = []
    for row in ROW_PATTERN.findall(html):
        cells = [clean_text(cell) for cell in CELL_PATTERN.findall(row)]
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

def display_results(results):
    print("\n"+"="*100)
    seen_tables = set()
    for result in results:
        # Skip True markers
        if result is True:
            continue
        # Check if result is a table (list of rows, where each row is a list)
        if isinstance(result, list) and len(result) > 0 and isinstance(result[0], list):
            # Use id() to track unique tables to avoid printing duplicates
            table_id = id(result)
            if table_id not in seen_tables:
                seen_tables.add(table_id)
                print("\n[TABLE]")
                for row in result:
                    # Format row as a table
                    print(" | ".join(str(cell) for cell in row))
                print("-"*100)
        elif isinstance(result, dict):
            # If it's a section dict, extract and display tables and paragraphs
            if "tables" in result:
                for table in result.get("tables", []):
                    table_id = id(table)
                    if table_id not in seen_tables:
                        seen_tables.add(table_id)
                        print("\n[TABLE]")
                        for row in table:
                            print(" | ".join(str(cell) for cell in row))
                        print("-"*100)
            # Also print paragraphs
            for paragraph in result.get("paragraphs", []):
                for sentence in paragraph:
                    if isinstance(sentence, str) and sentence.strip():
                        print(sentence)
                        print("-"*100)
        else:
            if result:  # Only print non-empty results
                print(result)
                print("-"*100)

def get_report(rcept_no):
    import os
    filename = f"{rcept_no}.json"
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        raw_text = unpack_zip(get_dart_report(rcept_no))
        data = extract_sections(raw_text)
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return data

if __name__ == "__main__":
    data = get_report("20251103000190")

    section = search(data, "기타위험", 0)
    texts = search(section, "희석", 0)

    section = search(data, "공모개요", 0)
    first_table = search_tables(section, "증권수량", 2)
    stock_ipo_amount = first_table[0][1][1]
    print(stock_ipo_amount)
    second_table = search_tables(section, "인수인", 2)
    print(second_table)
    amount = search_tables(second_table, "인수수량", 2)
    print(amount)

    section = search(data, "공모방법", 0)
    text = search(section, "신주모집")
    print(text)