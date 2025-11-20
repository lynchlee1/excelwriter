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

def search(node, keyword, parent_count=0, parent_node=None):
    """
    1) Recursively search `node` for `keyword`
    2) Return the ancestor `parent_count` levels above the matching node (0 for the matching node itself)
    """
    # parent_node is always a list of nodes
    if parent_node is None: parents = []
    elif isinstance(parent_node, list): parents = parent_node
    else: parents = [parent_node]

    matches = []
    def add_match(current_node, ancestry):
        if parent_count == 0: matches.append(current_node)
        elif parent_count <= len(ancestry): matches.append(ancestry[-parent_count])

    if isinstance(node, dict):
        for key, value in node.items():
            if isinstance(key, str) and keyword in key: add_match(value, parents + [node])
            matches.extend(search(value, keyword, parent_count, parents + [node]))
    elif isinstance(node, list):
        for item in node: matches.extend(search(item, keyword, parent_count, parents + [node]))
    elif isinstance(node, str):
        if keyword in node: add_match(node, parents)
    return matches

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
    
    # Search for "기타위험" first, then search those results for "희석"
    print("=== Searching for '기타위험' then '희석' ===")
    results = search(data, "기타위험", 0)
    # Now search the results for the second keyword
    # If results contain tables or strings, search through them
    filtered_results = []
    for result in results:
        # Search this result for the second keyword
        sub_results = search(result, "희석", 0)
        if sub_results:
            filtered_results.extend(sub_results)
        # Also check if the result itself contains both keywords
        elif isinstance(result, str) and "희석" in result:
            filtered_results.append(result)
        elif isinstance(result, list) and len(result) > 0:
            # Check if it's a table and contains the keyword
            if isinstance(result[0], list):
                # It's a table, check all cells
                for row in result:
                    for cell in row:
                        if isinstance(cell, str) and "희석" in cell:
                            filtered_results.append(result)
                            break
                    if result in filtered_results:
                        break
    
    if filtered_results:
        display_results(filtered_results)
    else:
        # If no results from filtered search, show results from first search
        display_results(results)

    # Second set of searches
    print("\n=== Searching for '공모개요' then '청약기일' ===")
    results = search(data, "공모개요", 0)
    filtered_results = []
    for result in results:
        sub_results = search(result, "청약기일", 0)
        if sub_results:
            filtered_results.extend(sub_results)
        elif isinstance(result, str) and "청약기일" in result:
            filtered_results.append(result)
        elif isinstance(result, list) and len(result) > 0:
            if isinstance(result[0], list):
                for row in result:
                    for cell in row:
                        if isinstance(cell, str) and "청약기일" in cell:
                            filtered_results.append(result)
                            break
                    if result in filtered_results:
                        break
    
    if filtered_results:
        display_results(filtered_results)
    else:
        display_results(results)
