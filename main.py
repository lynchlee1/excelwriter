import re
from opendart import get_report, get_dart_report
from webDataParser import number_value, unpack_zip, search_sections, search_tables, search, lookup_column, lookup_row, lookup_cell
from config import API_KEY

def parse_date(date_text, format):
    date_text = date_text.strip()
    match = re.match(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', date_text)
    if match:
        year, month, day = match.groups()
        month = month.zfill(2)
        day = day.zfill(2)
        if format == "yyyy.mm.dd": return f"{year}.{month}.{day}"
        elif format == "yy.mm.dd": return f"{year[-2:]}.{month}.{day}"
        elif format == "mm.dd": return f"{month}.{day}"
    return date_text


def display(header, data, apply_function=None):
    if apply_function: data = apply_function(data)
    if isinstance(data, list):
        for i in range(len(data)): 
            print(f"{header} {i+1}:", data[i])
    elif isinstance(data, dict):
        for key, value in data.items(): display(f"{header} {key}:", value)
    else: print(f"{header}: {data}")

from excelwriter import ExcelFile
if __name__ == "__main__":
    rcept_no_list = [
        "20251121000355", "20251110000199", "20251107000522",
    ]
    excel_file = f"ipo_reports.xlsx"
    excel_writer = ExcelFile(excel_file)
    for rcept_no in rcept_no_list:
        print("="*100)
        raw_text = unpack_zip(get_dart_report(rcept_no, API_KEY))
        with open(f"{rcept_no}_original.txt", "w", encoding="utf-8") as f: f.write(raw_text)
        data = get_report(rcept_no, API_KEY)
        sheet_name = f"{rcept_no}"
        excel_writer.clear_sheet(sheet_name)

        section_risk = search_sections(data, "기타위험")
        stock_option_texts = search(section_risk, include_keywords="희석", parent_count=0)
        new_stock_option_texts = []
        for text in stock_option_texts:
            if "상장" in text: new_stock_option_texts.append(text)
        excel_writer.write_cell(sheet_name, (1, 1), "희석")
        excel_writer.write_table_list(sheet_name, (1, 2), [new_stock_option_texts])

        section_ipo = search_sections(data, "공모개요")
        ipo_table_1 = search_tables(section_ipo, parent_count=2, include_keywords="증권수량")
        ipo_stock_count = lookup_column(ipo_table_1[0], "증권수량")
        excel_writer.write_cell(sheet_name, (2, 1), "증권수량")
        excel_writer.write_cell(sheet_name, (2, 2), ipo_stock_count)

        ipo_table_2 = search_tables(section_ipo, parent_count=2, include_keywords="인수인")
        ipo_holder_names = lookup_column(ipo_table_2[0], "인수인")
        ipo_holder_amounts = lookup_column(ipo_table_2[0], "인수금액")
        display("인수인", ipo_holder_names)
        display("인수규모", number_value(ipo_holder_amounts, 8, unit="억"))

        ipo_table_3 = search_tables(section_ipo, parent_count=2, include_keywords="청약기일")
        application_date = lookup_column(ipo_table_3[0], "청약기일")
        format_application = lambda value: [parse_date(item, format="yy.mm.dd") for item in value] if isinstance(value, list) else parse_date(value, format="yy.mm.dd")
        display("청약기일", application_date, format_application)
        payment_date = lookup_column(ipo_table_3[0], "납입기일")
        display("납입기일", payment_date, format_application)


        section_2 = search_sections(data, "공모방법")
        table_2_1 = search_tables(section_2, parent_count=2, include_keywords="기관투자자")
        institutional_pie = lookup_row(table_2_1[0], "기관투자자")
        display("기관투자자", institutional_pie)

        section_3 = search_sections(data, "공모가격결정방법")
        table_3_1 = search_tables(section_3, include_keywords=["희망공모","공모희망"], parent_count=2)
        price_range = lookup_row(table_3_1[0], ["희망공모","공모희망"])
        display("공모가 범위", price_range)
    
    excel_writer.save()
