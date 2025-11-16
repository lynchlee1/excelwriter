import os
import sys
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

class ExcelFile:
    def __init__(self, filename: str, basic_style: dict={}):
        '''
        Initialize object with filename.
        Reads local excel file in same directory. Creates new one if it doesn't exist.
        '''
        self.file_path = self.get_resource_path(filename)
        if os.path.exists(self.file_path): self.workbook = load_workbook(self.file_path) # Load existing workbook or create new one
        else: # Create new workbook
            self.workbook = Workbook()
            if 'Sheet' in self.workbook.sheetnames: self.workbook.remove(self.workbook['Sheet']) # Remove default sheet

    def get_resource_path(self, relative_path: str):
        '''
        Get current path, checking all possible scenarios.
        '''
        if getattr(sys, 'frozen', False): # Check if running as compiled .exe
            exe_dir = os.path.dirname(sys.executable)
            file_path = os.path.join(exe_dir, relative_path)
            if os.path.exists(file_path): return file_path
            try: base_path = sys._MEIPASS # Fall back to PyInstaller temp directory
            except AttributeError: base_path = exe_dir
        else: # Running as script
            current_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = current_dir
        return os.path.join(base_path, relative_path)
    
    def save(self):
        '''
        Save workbook to file.
        '''
        self.workbook.save(self.file_path)
    
    def get_sheet(self, sheet: str):
        '''
        Get existing sheet or create new one if it doesn't exist.
        '''
        if sheet in self.workbook.sheetnames: return self.workbook[sheet]
        else: return self.workbook.create_sheet(title=sheet)
    
    def clear_sheet(self, sheet: str):
        '''
        Clears sheet.
        '''
        ws = self.get_sheet(sheet)
        if ws.max_row > 0: ws.delete_rows(1, ws.max_row)
    
    def write_rows(self, sheet: str="", position: tuple=(1,1), datas: list=[], headers: list=[], show_headers: bool = True, header_style: dict = {}, content_style: dict = {}):
        '''
        Write table data to sheet.
        position: tuple of (start_row, start_col)
        datas: list of dict, whereach dict is a row of data
        headers: list of column headers
        header_style: style dict for header cells (overrides basic_style)
        content_style: style dict for content cells (overrides basic_style)
        '''
        ws = self.get_sheet(sheet)
        start_row, start_col = position
        
        # for key, value in enumerate(list, int) -> (int, list[0]), (int+1, list[1]), (int+2, list[2]), ...
        row = start_row
        if show_headers:
            for col, header in enumerate(headers, start_col):
                ws.cell(row, col).value = header
                if header_style: self._apply_cell_style(ws.cell(row, col), **header_style)
            row += 1
        for data in datas:
            for col, header in enumerate(headers, start_col):
                ws.cell(row, col).value = data.get(header)
                if content_style: self._apply_cell_style(ws.cell(row, col), **content_style)
            row += 1
    
    def write_cols(self, sheet: str="", position: tuple = (1, 1), datas: dict={}, headers: list=[], show_headers: bool = True, header_style: dict = {}, content_style: dict = {}):
        '''
        Write table data to sheet.
        position: tuple of (start_row, start_col)
        datas: dict of (header, list) pairs, where each list is a column of data
        headers: list of column headers
        header_style: style dict for header cells
        content_style: style dict for content cells
        '''
        ws = self.get_sheet(sheet)
        start_row, start_col = position

        # for key, value in enumerate(list, int) -> (int, list[0]), (int+1, list[1]), (int+2, list[2]), ...
        data_start_row = start_row + 1 if show_headers else start_row
        for col, header in enumerate(headers, start_col):
            if show_headers:
                ws.cell(start_row, col).value = header
                if header_style: self._apply_cell_style(ws.cell(start_row, col), **header_style)
            for row, value in enumerate(datas.get(header, []), data_start_row):
                ws.cell(row, col).value = value
                if content_style: self._apply_cell_style(ws.cell(row, col), **content_style)
    
    def write_cell(self, sheet: str="", position: tuple=(1,1), content: any=None, content_style: dict = {}):
        '''
        Write content to cell.
        position: tuple of (row, col)
        content: content to write to cell
        content_style: style dict for content cell (overrides basic_style)
        '''
        ws = self.get_sheet(sheet)
        row, col = position
        ws.cell(row, col).value = content
        if content_style: self._apply_cell_style(ws.cell(row, col), **content_style)
    
    def _apply_cell_style(self, cell, **kwargs):
        '''
        Apply cell style to cell.
        kwargs: keyword arguments for cell style
        '''
        if 'font' in kwargs:
            cell.font = kwargs.pop('font')
        elif font_props := {k: kwargs.pop(k) for k in ['name', 'size', 'bold', 'italic', 'underline', 'color'] if k in kwargs}:
            cell.font = Font(**font_props)
        
        if 'fill' in kwargs:
            cell.fill = kwargs.pop('fill')
        elif fill_props := {k: kwargs.pop(k) for k in ['fgColor', 'bgColor', 'fill_type', 'patternType'] if k in kwargs}:
            if 'fgColor' in fill_props: fill_props['start_color'] = fill_props.pop('fgColor')
            if 'bgColor' in fill_props: fill_props['end_color'] = fill_props.pop('bgColor')
            if 'patternType' in fill_props: fill_props['fill_type'] = fill_props.pop('patternType')
            cell.fill = PatternFill(**fill_props)
        
        if 'alignment' in kwargs:
            cell.alignment = kwargs.pop('alignment')
        elif align_props := {k: kwargs.pop(k) for k in ['horizontal', 'vertical', 'wrapText'] if k in kwargs}:
            if 'wrap_text' in kwargs: align_props['wrapText'] = kwargs.pop('wrap_text')
            if align_props: cell.alignment = Alignment(**align_props)
        
        if 'border' in kwargs:
            cell.border = kwargs.pop('border')
        elif 'border_style' in kwargs or 'border_color' in kwargs:
            side = Side(style=kwargs.pop('border_style', None), color=kwargs.pop('border_color', None))
            cell.border = Border(left=side, right=side, top=side, bottom=side)
        elif border_sides := {side: kwargs.pop(side) for side in ['left', 'right', 'top', 'bottom'] if side in kwargs}:
            for side in ['left', 'right', 'top', 'bottom']:
                if side not in border_sides:
                    border_sides[side] = getattr(cell.border, side, None) if cell.border else Side()
            cell.border = Border(**border_sides)
        
        for attr in ['number_format', 'hyperlink', 'comment', 'protection']:
            if attr in kwargs: setattr(cell, attr, kwargs.pop(attr))
        for key, value in kwargs.items():
            if hasattr(cell, key): setattr(cell, key, value)

if __name__ == "__main__":
    excel_file = ExcelFile("test.xlsx")
    ws = excel_file.get_sheet("styling_examples")
    row = 1
    
    # Title
    excel_file.write_cell("styling_examples", (row, 1), "Complete Style Examples", content_style={'bold': True, 'size': 16, 'fgColor': '4472C4', 'color': 'FFFFFF', 'horizontal': 'center'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    row += 3
    
    # FONT STYLES
    excel_file.write_cell("styling_examples", (row, 1), "FONT PROPERTIES", content_style={'bold': True, 'size': 14, 'fgColor': 'D9E1F2'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    excel_file.write_cell("styling_examples", (row, 1), "Property", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 2), "Example", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 3), "Values", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 4), "Code", content_style={'bold': True})
    row += 1
    
    font_examples = [
        ("name", "Font Name", "Arial", {'name': 'Arial'}),
        ("size", "Font Size", "16pt", {'size': 16}),
        ("bold", "Bold Text", "Bold", {'bold': True}),
        ("italic", "Italic Text", "Italic", {'italic': True}),
        ("underline", "Underline", "Underlined", {'underline': 'single'}),
        ("color", "Font Color", "Red Text", {'color': 'FF0000'}),
    ]
    for prop, label, value, style in font_examples:
        excel_file.write_cell("styling_examples", (row, 1), prop, content_style={'italic': True})
        excel_file.write_cell("styling_examples", (row, 2), label, content_style=style)
        excel_file.write_cell("styling_examples", (row, 3), value)
        excel_file.write_cell("styling_examples", (row, 4), str(style))
        row += 1
    excel_file.write_cell("styling_examples", (row, 2), "Combined", content_style={'bold': True, 'italic': True, 'size': 14, 'color': '0000FF'})
    excel_file.write_cell("styling_examples", (row, 4), "{'bold': True, 'italic': True, 'size': 14, 'color': '0000FF'}")
    row += 2
    
    # FILL STYLES
    excel_file.write_cell("styling_examples", (row, 1), "FILL PROPERTIES", content_style={'bold': True, 'size': 14, 'fgColor': 'D9E1F2'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    excel_file.write_cell("styling_examples", (row, 1), "Property", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 2), "Example", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 3), "Color Code", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 4), "Code", content_style={'bold': True})
    row += 1
    
    fill_examples = [
        ("fgColor", "Yellow Background", "FFFF00", {'fgColor': 'FFFF00'}),
        ("fgColor", "Green Background", "00FF00", {'fgColor': '00FF00'}),
        ("fgColor", "Blue Background", "0000FF", {'fgColor': '0000FF', 'color': 'FFFFFF'}),
        ("fgColor", "Red Background", "FF0000", {'fgColor': 'FF0000', 'color': 'FFFFFF'}),
        ("fgColor", "Gray Background", "CCCCCC", {'fgColor': 'CCCCCC'}),
        ("fill_type", "Pattern Fill", "solid", {'fgColor': 'FFC7CE', 'fill_type': 'solid'}),
    ]
    for prop, label, code, style in fill_examples:
        excel_file.write_cell("styling_examples", (row, 1), prop, content_style={'italic': True})
        excel_file.write_cell("styling_examples", (row, 2), label, content_style=style)
        excel_file.write_cell("styling_examples", (row, 3), code)
        excel_file.write_cell("styling_examples", (row, 4), str(style))
        row += 1
    row += 2
    
    # ALIGNMENT STYLES
    excel_file.write_cell("styling_examples", (row, 1), "ALIGNMENT PROPERTIES", content_style={'bold': True, 'size': 14, 'fgColor': 'D9E1F2'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    excel_file.write_cell("styling_examples", (row, 1), "Property", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 2), "Example", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 3), "Values", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 4), "Code", content_style={'bold': True})
    row += 1
    
    align_examples = [
        ("horizontal", "Left Aligned", "left", {'horizontal': 'left', 'fgColor': 'E7E6E6'}),
        ("horizontal", "Center Aligned", "center", {'horizontal': 'center', 'fgColor': 'E7E6E6'}),
        ("horizontal", "Right Aligned", "right", {'horizontal': 'right', 'fgColor': 'E7E6E6'}),
        ("vertical", "Top Aligned", "top", {'vertical': 'top', 'fgColor': 'E7E6E6'}),
        ("vertical", "Middle Aligned", "center", {'vertical': 'center', 'fgColor': 'E7E6E6'}),
        ("vertical", "Bottom Aligned", "bottom", {'vertical': 'bottom', 'fgColor': 'E7E6E6'}),
        ("wrapText", "Wrap Text", "This text wraps to multiple lines when the cell width is narrow", {'wrapText': True, 'fgColor': 'E7E6E6'}),
    ]
    for prop, label, value, style in align_examples:
        excel_file.write_cell("styling_examples", (row, 1), prop, content_style={'italic': True})
        excel_file.write_cell("styling_examples", (row, 2), label, content_style=style)
        excel_file.write_cell("styling_examples", (row, 3), value)
        excel_file.write_cell("styling_examples", (row, 4), str(style))
        row += 1
    row += 2
    
    # BORDER STYLES
    excel_file.write_cell("styling_examples", (row, 1), "BORDER PROPERTIES", content_style={'bold': True, 'size': 14, 'fgColor': 'D9E1F2'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    excel_file.write_cell("styling_examples", (row, 1), "Property", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 2), "Example", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 3), "Style", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 4), "Code", content_style={'bold': True})
    row += 1
    
    border_examples = [
        ("border_style", "Thin Border", "thin", {'border_style': 'thin'}),
        ("border_style", "Thick Border", "thick", {'border_style': 'thick'}),
        ("border_style", "Dashed Border", "dashed", {'border_style': 'dashed'}),
        ("border_style", "Double Border", "double", {'border_style': 'double'}),
        ("border_color", "Red Border", "thin + red", {'border_style': 'thin', 'border_color': 'FF0000'}),
        ("border_style", "Dotted Border", "dotted", {'border_style': 'dotted'}),
    ]
    for prop, label, value, style in border_examples:
        excel_file.write_cell("styling_examples", (row, 1), prop, content_style={'italic': True})
        excel_file.write_cell("styling_examples", (row, 2), label, content_style=style)
        excel_file.write_cell("styling_examples", (row, 3), value)
        excel_file.write_cell("styling_examples", (row, 4), str(style))
        row += 1
    row += 2
    
    # NUMBER FORMAT
    excel_file.write_cell("styling_examples", (row, 1), "NUMBER FORMAT", content_style={'bold': True, 'size': 14, 'fgColor': 'D9E1F2'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    excel_file.write_cell("styling_examples", (row, 1), "Value", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 2), "Formatted", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 3), "Format Code", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 4), "Style Code", content_style={'bold': True})
    row += 1
    
    number_examples = [
        (1234.56, "#,##0.00", {'number_format': '#,##0.00'}),
        (0.15, "0.00%", {'number_format': '0.00%'}),
        (1234.56, "$#,##0.00", {'number_format': '$#,##0.00'}),
        (1234, "0", {'number_format': '0'}),
        (0.1234, "0.00", {'number_format': '0.00'}),
    ]
    for value, fmt, style in number_examples:
        excel_file.write_cell("styling_examples", (row, 1), value)
        excel_file.write_cell("styling_examples", (row, 2), value, content_style=style)
        excel_file.write_cell("styling_examples", (row, 3), fmt)
        excel_file.write_cell("styling_examples", (row, 4), str(style))
        row += 1
    row += 2
    
    # COMBINED EXAMPLES
    excel_file.write_cell("styling_examples", (row, 1), "COMBINED STYLES", content_style={'bold': True, 'size': 14, 'fgColor': 'D9E1F2'})
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 1
    excel_file.write_cell("styling_examples", (row, 1), "Description", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 2), "Example", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 3), "Combined Properties", content_style={'bold': True})
    excel_file.write_cell("styling_examples", (row, 4), "Code", content_style={'bold': True})
    row += 1
    
    combined_examples = [
        ("Bold + Yellow + Center", {'bold': True, 'fgColor': 'FFFF00', 'horizontal': 'center'}),
        ("Italic + Blue + Right", {'italic': True, 'color': '0000FF', 'horizontal': 'right'}),
        ("Bold + Green + Border", {'bold': True, 'fgColor': '00FF00', 'border_style': 'thin'}),
        ("Large + Red + Center", {'size': 16, 'color': 'FF0000', 'horizontal': 'center', 'vertical': 'center'}),
        ("All Styles Combined", {'bold': True, 'size': 14, 'fgColor': 'FFC7CE', 'color': '000000', 'horizontal': 'center', 'vertical': 'center', 'border_style': 'thick', 'border_color': 'FF0000'}),
    ]
    for desc, style in combined_examples:
        excel_file.write_cell("styling_examples", (row, 1), desc, content_style={'italic': True})
        excel_file.write_cell("styling_examples", (row, 2), desc, content_style=style)
        excel_file.write_cell("styling_examples", (row, 3), ", ".join(f"{k}={v}" for k, v in style.items()))
        excel_file.write_cell("styling_examples", (row, 4), str(style))
        row += 1
    
    # Auto-adjust column widths
    from openpyxl.utils import get_column_letter
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 25
    
    excel_file.save()
    print("Complete styling examples written to test.xlsx")