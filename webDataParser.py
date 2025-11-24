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


def number_value(obj: str | list, power: int=0, round_digits: int=1, unit: str=""):
    '''
    Convert string/list to float.
    If unit is provided, return string with unit.
    '''
    if isinstance(obj, list): 
        return [number_value(item, power, round_digits, unit) for item in obj]
    elif isinstance(obj, str): 
        obj = float(obj.replace(',', '').strip())
        if unit: return str(round(obj / 10**power, round_digits)) + unit
        else: return round(obj / 10**power, round_digits)

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
def search(
    node, parent_node=None, parent_count=0,
    include_keywords: str | list = None, 
    exclude_keywords: str | list = None,
    exact: bool = False
    ):
    '''
    1) Recursively search node that contains `include_keywords` and does not contain `exclude_keywords` (can be a string or list of strings for OR search)
    2) Return the ancestor `parent_count` levels above the matching node (0 for the matching node itself)
    3) exact: if True, match exactly instead of substring match
    '''
    if isinstance(include_keywords, str): include_keywords = [include_keywords]
    if isinstance(exclude_keywords, str): exclude_keywords = [exclude_keywords]
    
    def _match(text, keyword_list):
        if exact: return any(text == keyword for keyword in keyword_list)
        else: return any(keyword in text for keyword in keyword_list)
    
    def _check_match(text):
        if include_keywords and not _match(text, include_keywords): return False
        if exclude_keywords and _match(text, exclude_keywords): return False
        return True
    
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
            if isinstance(key, str) and _check_match(key): _add_match(value, parents + [node])
            kwargs = {
                "parent_node": parents + [node],
                "parent_count": parent_count,
                "include_keywords": include_keywords,
                "exclude_keywords": exclude_keywords,
                "exact": exact,
            }
            matches.extend(search(value, **kwargs))

    elif isinstance(node, list): 
        kwargs = {
            "parent_node": parents + [node],
            "parent_count": parent_count,
            "include_keywords": include_keywords,
            "exclude_keywords": exclude_keywords,
            "exact": exact,
        }
        for item in node: matches.extend(search(item, **kwargs))
    
    elif isinstance(node, str) and _check_match(node): _add_match(node, parents)
    
    return matches

def search_tables(
    sections, parent_count=0,
    include_keywords: str | list = None, 
    exclude_keywords: str | list = None,
    exact: bool = False
    ):
    '''
    1) Search tables in sections that contains `include_keywords` and do not contain `exclude_keywords` (can be a string or list of strings for OR search)
    2) Return the ancestor `parent_count` levels above the matching node (0 for the matching node itself)
    3) exact: if True, match exactly instead of substring match
    '''
    if isinstance(include_keywords, str): include_keywords = [include_keywords]
    if isinstance(exclude_keywords, str): exclude_keywords = [exclude_keywords]
    kwargs = {
        "parent_count": parent_count,
        "include_keywords": include_keywords, 
        "exclude_keywords": exclude_keywords, 
        "exact": exact,
    }    
    matches = []
    for section in sections:
        if isinstance(section, dict):
            for table in section.get("tables", []): matches.extend(search(table, parent_node=section, **kwargs))
        elif isinstance(section, list): matches.extend(search(section, parent_node=section, **kwargs))
    return matches

def search_sections(
    article, 
    include_keywords: str | list = None, 
    exclude_keywords: str | list = None,
    exact: bool = False
    ):
    '''
    1) Search sections in article that contains `include_keywords` and do not contain `exclude_keywords` (can be a string or list of strings for OR search)
    2) exact: if True, match exactly instead of substring match
    3) Only checks section names (keys), not recursive search
    '''
    if isinstance(include_keywords, str): include_keywords = [include_keywords]
    if isinstance(exclude_keywords, str): exclude_keywords = [exclude_keywords]
    
    def _match(text, keyword_list):
        if exact: return any(text == keyword for keyword in keyword_list)
        else: return any(keyword in text for keyword in keyword_list)
    
    def _check_match(text):
        if include_keywords and not _match(text, include_keywords): return False
        if exclude_keywords and _match(text, exclude_keywords): return False
        return True
    
    matches = []
    for section_name, section_content in article.items():
        if _check_match(section_name): matches.append(section_content)
    return matches

'''
Excel-like lookup functions
'''
def lookup_column(table, keys, exact=False):
    '''
    Given a table (list of lists, table[0] is header), 
    returns all row values in the key columns as a list of lists if multiple columns match, else single list.
    '''
    if isinstance(keys, str): keys = [keys]
    matches_idx = []
    col_len = len(table[0])
    for col_idx in range(col_len):
        if exact:
            if table[0][col_idx] in keys: matches_idx.append(col_idx)
        else:
            if any(key in table[0][col_idx] for key in keys): matches_idx.append(col_idx)
    if not matches_idx: return []

    result = []
    for col_idx in matches_idx:
        col_values = [table[row_idx][col_idx] for row_idx in range(1, len(table))]
        result.append(col_values)

    if len(result) == 1: 
        if len(result[0]) == 1: return result[0][0]
        else: return result[0]

    return result

def lookup_row(table, keys, exact=False):
    '''
    Given a table (list of lists, table[row][0] is header), 
    returns all column values in the key row as a list of lists if multiple rows match, else single list.
    '''
    if isinstance(keys, str): keys = [keys]
    matches_idx = []
    row_len = len(table)
    for row_idx in range(row_len):
        if exact:
            if table[row_idx][0] in keys: matches_idx.append(row_idx)
        else:
            if any(key in table[row_idx][0] for key in keys): matches_idx.append(row_idx)
    if not matches_idx: return []

    result = []
    for row_idx in matches_idx:
        row_values = [table[row_idx][col_idx] for col_idx in range(1, len(table[row_idx]))]
        result.append(row_values)

    if len(result) == 1: 
        if len(result[0]) == 1: return result[0][0]
        else: return result[0]

    return result

def lookup_cell(table, row_keys, col_keys, row_exact=False, col_exact=False):
    '''
    Given a table (list of lists, table[row][col] is cell), 
    returns the cell value if the row and column keys match, else None.
    '''
    if isinstance(row_keys, str): row_keys = [row_keys]
    if isinstance(col_keys, str): col_keys = [col_keys]
    row_idx_found = []
    col_idx_found = []

    for row_idx in range(1, len(table)):
        cell_value = table[row_idx][0]
        if row_exact:
            if cell_value in row_keys: row_idx_found.append(row_idx)
        else:
            if any(key in cell_value for key in row_keys): row_idx_found.append(row_idx)
    if not row_idx_found: return None

    for col_idx in range(1, len(table[0])):
        col_value = table[0][col_idx]
        if col_exact:
            if col_value in col_keys: col_idx_found.append(col_idx)
        else:
            if any(key in col_value for key in col_keys): col_idx_found.append(col_idx)
    if not col_idx_found: return None

    result = []
    for row_idx in row_idx_found:
        for col_idx in col_idx_found:
            result.append(table[row_idx][col_idx])
    
    if len(result) == 1: 
        if len(result[0]) == 1: return result[0][0]
        else: return result[0]
    
    return result


'''
I dont understand why this works...
'''

def parse_table(html):
    '''
    Get html table and return list of rows. 
    Handles colspan/rowspan properly.
    '''
    rows, rowspans = [], []
    for row_html in ROW_PATTERN.findall(html): # for each <tr></tr> block
        cells = []
        # Track which columns have been filled by rowspans from previous rows
        filled_by_rowspan = set()
        # First, handle rowspans from previous rows by inserting them at their column positions
        for col_idx in range(len(rowspans)):
            if rowspans[col_idx]:
                r, text = rowspans[col_idx]
                # Ensure cells list is long enough
                while len(cells) <= col_idx:
                    cells.append('')
                cells[col_idx] = text
                filled_by_rowspan.add(col_idx)  # Mark this column as filled by rowspan
                # Decrement the rowspan counter
                rowspans[col_idx] = (r - 1, text) if r > 1 else None
            else:
                # Ensure cells list is long enough even for None rowspans
                while len(cells) <= col_idx:
                    cells.append('')
        
        col = 0  # Current column position in the row

        for m in CELL_PATTERN.finditer(row_html): # for each <td></td> or <th></th> block
            # Skip columns that have active rowspans (from previous rows) or are already filled by rowspan
            while col < len(rowspans) and (rowspans[col] is not None or col in filled_by_rowspan):
                col += 1
            
            text = clean_text(m.group(1)) # get text content of the cell
            tag = CELL_TAG_PATTERN.search(row_html, m.start()).group(0) # get tag of the cell

            # Get colspan/rowspan values from the tag
            m_colspan = re.search(r'colspan\s*=\s*["\']?(\d+)', tag, re.I)
            colspan = int(m_colspan.group(1)) if m_colspan else 1 
            m_rowspan = re.search(r'rowspan\s*=\s*["\']?(\d+)', tag, re.I)
            rowspan = int(m_rowspan.group(1)) if m_rowspan else 1

            # Handle rowspan: track it for all columns spanned by colspan
            if rowspan > 1:
                for i in range(colspan):
                    col_idx = col + i
                    # Ensure rowspans list is long enough
                    while len(rowspans) <= col_idx:
                        rowspans.append(None)
                    rowspans[col_idx] = (rowspan - 1, text)
            
            for i in range(colspan): # for each colspan
                # Ensure cells list is long enough
                while len(cells) <= col:
                    cells.append('')
                # Only set text if cell is empty (don't overwrite rowspanned cells)
                if cells[col] == '':
                    cells[col] = text
                col += 1
        
        # Clean up trailing None entries in rowspans
        while rowspans and rowspans[-1] is None:
            rowspans.pop()
        
        if cells:
            rows.append(cells) # add row to the table

    if rows:
        # Append empty cells to match the number of columns
        max_cols = max(len(r) for r in rows)
        return [r + [''] * (max_cols - len(r)) for r in rows]

    return rows
