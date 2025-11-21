import re
import zipfile
from io import BytesIO

'''
Regular expressions 
'''
TABLE_PATTERN = re.compile(r'<TABLE[^>]*>.*?</TABLE>', re.DOTALL | re.IGNORECASE)
ROW_PATTERN = re.compile(r'<TR[^>]*>(.*?)</TR>', re.DOTALL | re.IGNORECASE)
CELL_PATTERN = re.compile(r'<T[DH][^>]*>(.*?)</T[DH]>', re.DOTALL | re.IGNORECASE)


'''
HTML handling functions
'''
def clean_text(text): # clean HTML tags and extra spaces
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', text)).strip()


'''
Unpack web data to text
'''
def unpack_zip(data):
    with zipfile.ZipFile(BytesIO(data)) as zf: content = zf.read(zf.namelist()[0])
    try: return content.decode('utf-8')
    except UnicodeDecodeError: return content.decode('cp949', errors='ignore')


'''
Search functions 
'''
def search(node, keyword, parent_count=0, parent_node=None):
    """
    1) Recursively search node that contains `keyword`
    2) Return the ancestor `parent_count` levels above the matching node (0 for the matching node itself)
    """
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
    tables = []
    for section in sections:
        if isinstance(section, dict):
            for table in section.get("tables", []): tables.extend(search(table, keyword, parent_count))
        elif isinstance(section, list): tables.extend(search(section, keyword, parent_count))
    return tables
