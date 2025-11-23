import re
from opendart import get_report, get_dart_report
from webDataParser import sequential_search, number_value, unpack_zip
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


def lookup_column(table, key):
    '''
    Given a table (list of lists, table[0] is header), 
    returns all row values in the key columns as a list of lists if multiple columns match, else single list.
    '''
    matches_idx = []
    col_len = len(table[0])
    for col_idx in range(col_len):
        if table[0][col_idx] == key: matches_idx.append(col_idx)
    if not matches_idx: return []

    result = []
    for col_idx in matches_idx:
        col_values = [table[row_idx][col_idx] for row_idx in range(1, len(table))]
        result.append(col_values)

    return result


if __name__ == "__main__":
    rcept_no = "20251024000492"
    raw_text = unpack_zip(get_dart_report(rcept_no, API_KEY))
    with open(f"{rcept_no}_original.txt", "w", encoding="utf-8") as f:
        f.write(raw_text)
    data = get_report(rcept_no, API_KEY)

    stock_option_texts = sequential_search(data, [("기타위험", 0), ("희석", 0)])
    print("희석주식수 관련 텍스트: ", stock_option_texts)

    ipo_table_1 = sequential_search(data, [("공모개요", 0), ("증권수량", 2)])[0]
    ipo_stock_count = lookup_column(ipo_table_1, "증권수량")
    print("증권수량: ", ipo_stock_count)

    ipo_table_2 = sequential_search(data, [("공모개요", 0), ("인수인", 2)])[0]
    ipo_holder_count = lookup_column(ipo_table_2, "인수인")
    print("인수인: ", ipo_holder_count)
    # holder_texts = []
    # for row in ipo_table_2:
    #     try: holder_texts.append(row[1]+" "+str(round(number_value(row[4])/10**8, 1))+"억")
    #     except: pass

    ipo_table_3 = sequential_search(data, [("공모개요", 0), ("청약기일", 2)])[0]
    print(ipo_table_3)
    application_date_idx = ipo_table_3[0].index("청약기일")
    application_date = ipo_table_3[1][application_date_idx]
    application_date_parts = [part.strip() for part in application_date.split("~")]

    # print("청약기일 분리: ", parse_date(application_date_parts[0],"yy.mm.dd"), parse_date(application_date_parts[1],"mm.dd"))

    payment_date_idx = ipo_table_3[0].index("납입기일")
    payment_date = ipo_table_3[1][payment_date_idx]

