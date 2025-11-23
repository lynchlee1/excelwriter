import os
import json
import requests
from webDataParser import unpack_zip, extract_sections
from excelwriter import ExcelFile

'''
개별 공시 가져오기
'''
def get_dart_report(rcept_no: str, api_key: str):
    url = f"https://opendart.fss.or.kr/api/document.xml?crtfc_key={api_key}&rcept_no={rcept_no}"
    response = requests.get(url)
    return response.content

def get_report(rcept_no: str, api_key: str = ""):
    filename = f"{rcept_no}.json"
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        if not api_key:
            raise ValueError("api_key is required when report file doesn't exist")
        raw_text = unpack_zip(get_dart_report(rcept_no, api_key))
        data = extract_sections(raw_text)
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    return data

'''
IPO 관련 공시 가져오기
'''
def get_all_ipo_reports(start_date: str, end_date: str, api_key: str="", save_path: str="", last_reprt_at: str="N"):
    '''
    Save all IPO reports between start_date and end_date to save_path.
    last_reprt_at: Y for latest reports, N for all reports.
    '''
    all_items = []
    page_no = 1

    url = f"https://opendart.fss.or.kr/api/list.json?crtfc_key={api_key}&bgn_de={start_date}&end_de={end_date}&last_reprt_at={last_reprt_at}&pblntf_ty=C&pblntf_detail_ty=C001&page_no={page_no}&page_count=100"
    response = requests.get(url)
    first_data = response.json()
    if first_data.get("status") != "000": return first_data # when first request is invalid
    
    all_items = []
    all_items.extend(first_data.get("list"))
    total_page = first_data.get("total_page", 1)
    for page_no in range(2, total_page + 1): # when total_page > 1, request additional pages
        url = f"https://opendart.fss.or.kr/api/list.json?crtfc_key={api_key}&bgn_de={start_date}&end_de={end_date}&last_reprt_at={last_reprt_at}&pblntf_ty=C&pblntf_detail_ty=C001&page_no={page_no}&page_count=100"
        response = requests.get(url)
        page_data = response.json()
        if page_data.get("status") == "000": 
            print(page_data.get("message")) # "정상" is printed when request is valid
            all_items.extend(page_data.get("list"))
    
    combined_data = {
        "status": first_data.get("status"),
        "total_count": first_data.get("total_count"),
        "list": all_items
    }

    if save_path:
        with open(save_path, 'w', encoding='utf-8') as f: json.dump(combined_data, f, ensure_ascii=False, indent=2)
    return combined_data


def get_all_ipo_reports_multi_year(start_year: int, end_year: int, api_key: str="", save_path: str="", folder_name: str="", last_reprt_at: str="N"):
    '''
    Collects IPO reports across multiple years by iterating through quarterly periods.
    Saves each quarter's data to a separate JSON file to prevent corruption.
    '''
    all_items = []
    total_count = 0
    
    # (start_month, end_month, start_day, end_day)
    quarters = [
        (1, 3, 1, 31),   # Q1: Jan 01 - Mar 31
        (4, 6, 1, 30),   # Q2: Apr 01 - Jun 30
        (7, 9, 1, 30),   # Q3: Jul 01 - Sep 30
        (10, 12, 1, 31), # Q4: Oct 01 - Dec 31
    ]
    
    for year in range(start_year, end_year + 1):
        for start_month, end_month, start_day, end_day in quarters:
            start_date = f"{year}{start_month:02d}{start_day:02d}"
            end_date = f"{year}{end_month:02d}{end_day:02d}"
            
            print(f"Fetching data for {start_date} to {end_date}...")
            quarter_data = get_all_ipo_reports(start_date, end_date, api_key=api_key, last_reprt_at=last_reprt_at)
            
            if quarter_data.get("status") == "000":
                quarter_items = quarter_data.get("list")
                all_items.extend(quarter_items) 
                total_count += len(quarter_items)
                print(f"  Collected {len(quarter_items)} items (total: {total_count})")
                
                if folder_name:
                    os.makedirs(folder_name, exist_ok=True)
                    quarterly_file = os.path.join(folder_name, f"ipo_reports_{start_date}_{end_date}.json")
                    with open(quarterly_file, 'w', encoding='utf-8') as f:
                        json.dump(quarter_data, f, ensure_ascii=False, indent=2)
                    print(f"  Saved quarterly data to {quarterly_file}")
            else:
                print(f"  Error: {quarter_data.get('message', 'Unknown error')}")
    
    combined_data = {
        "total_count": total_count,
        "list": all_items
    }
    
    if save_path:
        with open(save_path, 'w', encoding='utf-8') as f: json.dump(combined_data, f, ensure_ascii=False, indent=2)
        print(f"\nSaved {total_count} total items to {save_path}")
    
    return combined_data

'''
Filter and process IPO reports
'''
def filter_json(filename: str, from_dir: str, to_dir: str):
    os.makedirs(to_dir, exist_ok=True)
    from_path, to_path = os.path.join(from_dir, filename), os.path.join(to_dir, filename)
    with open(from_path, "r", encoding="utf-8") as f: data = json.load(f)

    filtered_list = []
    for item in data.get("list", []):
        report_nm = item.get("report_nm", "")
        if "증권신고서(지분증권)" in report_nm: filtered_list.append(item)

    filtered_data = {
        "total_count": len(filtered_list),
        "list": filtered_list
    }
    with open(to_path, "w", encoding="utf-8") as f: json.dump(filtered_data, f, ensure_ascii=False, indent=2)
    return filtered_data

def filter_all_json_files(from_dir: str, to_dir: str):
    if not os.path.exists(from_dir): return None
    json_files = [f for f in os.listdir(from_dir) if f.endswith('.json')]
    json_files.sort()
    for filename in json_files:
        try: filter_json(filename, from_dir, to_dir)
        except Exception: return None

def concat_json_files(from_dir: str, output_file: str):
    if not os.path.exists(from_dir): return None
    json_files = [f for f in os.listdir(from_dir) if f.endswith('.json')]
    json_files.sort()
    all_items = []
    for filename in json_files:
        file_path = os.path.join(from_dir, filename)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                items = data.get("list", [])
                all_items.extend(items)
        except Exception: return None
    combined_data = {
        "total_count": len(all_items),
        "list": all_items
    }
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(combined_data, f, ensure_ascii=False, indent=2)
    return combined_data

def save_json_to_excel(json_file: str, excel_file: str):
    with open(json_file, "r", encoding="utf-8") as f: data = json.load(f)
    items = data.get("list", [])
    if not items: return None
    headers = []
    if items: headers = list(items[0].keys())
    excel = ExcelFile(excel_file)
    sheet_name = "IPO Reports"
    excel.clear_sheet(sheet_name)
    
    excel.write_rows(
        sheet=sheet_name,
        position=(1, 1),
        datas=items,
        headers=headers,
        show_headers=True
    )
    excel.save()
