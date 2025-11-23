import re
import zipfile
from io import BytesIO

'''
Regular expressions 
'''
TABLE_PATTERN = re.compile(r'<TABLE[^>]*>.*?</TABLE>', re.DOTALL | re.IGNORECASE)
ROW_PATTERN = re.compile(r'<TR[^>]*>(.*?)</TR>', re.DOTALL | re.IGNORECASE)
CELL_PATTERN = re.compile(r'<T[DH][^>]*>(.*?)</T[DH]>', re.DOTALL | re.IGNORECASE)
CELL_TAG_PATTERN = re.compile(r'<T[DH][^>]*>', re.IGNORECASE)

def split_paragraphs(text):
    '''
    Split text into paragraphs. 
    '''
    blocks = []
    for part in [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]:
        cleaned = clean_text(part)
        parts = [chunk.strip() for chunk in re.split(r'(?<=[.?!])\s+', cleaned) if chunk.strip()]
        if parts: blocks.append(parts)
    return blocks


def split_texts(text):
    '''
    Split text into paragraphs and tables.
    '''
    paragraphs = []
    tables = []
    last = 0
    for match in TABLE_PATTERN.finditer(text):
        before = text[last:match.start()]
        if before.strip(): paragraphs.extend(split_paragraphs(before))
        table_rows = parse_table(match.group(0))
        if table_rows: tables.append(table_rows)
        last = match.end()
    tail = text[last:]
    if tail.strip(): paragraphs.extend(split_paragraphs(tail))
    return {"paragraphs": paragraphs, "tables": tables}


def extract_sections(text):
    '''
    Extract sections from text.
    '''
    matches = list(re.finditer(r'<TITLE[^>]*>(.*?)</TITLE>', text, re.DOTALL | re.IGNORECASE))
    if not matches: return None
    sections = {}
    for i, match in enumerate(matches):
        title = re.sub(r'<[^>]+>', '', match.group(1)).strip()
        if title:
            end = matches[i+1].start() if i+1 < len(matches) else len(text)
            sections[title.replace(' ', '')] = split_texts(text[match.end():end].strip())
    return sections


def clean_text(text):
    '''
    Clean HTML tags and extra spaces.
    '''
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', text)).strip()


def number_value(s):
    '''
    Convert string to float.
    '''
    s = s.replace(',', '').strip()
    try: return float(s)
    except: return None


'''
Unpack web data to text
'''
def unpack_zip(data):
    '''
    Unpack zip file to text.
    '''
    with zipfile.ZipFile(BytesIO(data)) as zf: content = zf.read(zf.namelist()[0])
    try: return content.decode('utf-8')
    except UnicodeDecodeError: return content.decode('cp949', errors='ignore')


'''
Search functions 
'''
def search(node, keyword, parent_count=0, parent_node=None):
    '''
    1) Recursively search node that contains `keyword`
    2) Return the ancestor `parent_count` levels above the matching node (0 for the matching node itself)
    '''
    # parent_node is always a list of nodes
    if parent_node is None: parents = []
    elif isinstance(parent_node, list): parents = parent_node
    else: parents = [parent_node]

    matches = []
    def _add_match(current_node, ancestry):
        if parent_count == 0: matches.append(current_node) # add itself
        elif parent_count <= len(ancestry): matches.append(ancestry[-parent_count]) # add ancestor node

    if isinstance(node, dict):
        for key, value in node.items():
            if isinstance(key, str) and keyword in key: _add_match(value, parents + [node])
            matches.extend(search(value, keyword, parent_count, parents + [node]))

    elif isinstance(node, list): 
        for item in node: matches.extend(search(item, keyword, parent_count, parents + [node]))
    
    elif isinstance(node, str) and keyword in node: _add_match(node, parents)
    
    return matches

def search_tables(sections, keyword, parent_count=0):
    '''
    Search tables in sections.
    '''
    tables = []
    for section in sections:
        if isinstance(section, dict):
            for table in section.get("tables", []): tables.extend(search(table, keyword, parent_count))
        elif isinstance(section, list): tables.extend(search(section, keyword, parent_count))
    return tables

def sequential_search(node, keywords):
    '''
    Sequentially search for keywords in the node. Get list of (keyword, parent_count) pairs as keywords input.
    '''
    for keyword, parent_count in keywords: node = search(node, keyword, parent_count)
    return node


def parse_table(html):
    '''
    Get html table and return list of rows. 
    Considers colspan and rowspan.
    '''
    rows, rowspans = [], []
    for row_html in ROW_PATTERN.findall(html):
        cells = []
        col = 0
        while col < len(rowspans):
            if rowspans[col]:
                r, text = rowspans[col]
                cells.append(text)
                rowspans[col] = (r - 1, text) if r > 1 else None
            col += 1
        for m in CELL_PATTERN.finditer(row_html):
            while col < len(rowspans) and rowspans[col]:
                r, text = rowspans[col]
                cells.append(text)
                rowspans[col] = (r - 1, text) if r > 1 else None
                col += 1
            text = clean_text(m.group(1))
            tag = CELL_TAG_PATTERN.search(row_html, m.start()).group(0)
            m_colspan = re.search(r'colspan\s*=\s*["\']?(\d+)', tag, re.I)
            colspan = int(m_colspan.group(1)) if m_colspan else 1
            m_rowspan = re.search(r'rowspan\s*=\s*["\']?(\d+)', tag, re.I)
            rowspan = int(m_rowspan.group(1)) if m_rowspan else 1
            for _ in range(colspan):
                cells.append(text)
                if rowspan > 1:
                    while len(rowspans) <= col: rowspans.append(None)
                    rowspans[col] = (rowspan - 1, text)
                col += 1
        while rowspans and rowspans[-1] is None: rowspans.pop()
        if cells: rows.append(cells)
    if rows:
        max_cols = max(len(r) for r in rows)
        return [r + [''] * (max_cols - len(r)) for r in rows]
    return rows
