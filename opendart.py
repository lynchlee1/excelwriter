import requests
import zipfile
from io import BytesIO
import re
import json
from config import API_KEY

def decode_text(data):
    try: return data.decode('utf-8', errors='ignore')
    except UnicodeDecodeError:
        try: return data.decode('cp949', errors='ignore')
        except UnicodeDecodeError: return None

def extract_sections(text):
    sections = {}
    pattern = r'<TITLE[^>]*>(.*?)</TITLE>'
    matches = list(re.finditer(pattern, text, re.DOTALL | re.IGNORECASE))
    
    for i, match in enumerate(matches):
        title_text = re.sub(r'<[^>]+>', '', match.group(1)).strip()
        if title_text:
            start_pos = match.end()
            end_pos = matches[i+1].start() if i+1 < len(matches) else len(text)
            content = text[start_pos:end_pos].strip()
            sections[title_text] = content
    
    return sections if sections else None

def process_file(zip_file, filename):
    with zip_file.open(filename) as f:
        data = f.read()
        text = decode_text(data)
        if not text: return None
        sections = extract_sections(text)
        return sections if sections else text

def unpack(rcept_no: str):
    url = f"https://opendart.fss.or.kr/api/document.xml?crtfc_key={API_KEY}&rcept_no={rcept_no}"
    response = requests.get(url)
    result = {}
    try:
        with zipfile.ZipFile(BytesIO(response.content)) as zf:
            for filename in zf.namelist():
                content = process_file(zf, filename)
                if content: result[filename] = content
    except zipfile.BadZipFile: pass
    return result

data = unpack("20251024000492")
with open(f"opendart_20251024000492.json", 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)