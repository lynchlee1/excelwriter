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
        self.basic_style = {}

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
                if self.basic_style or header_style: self._apply_cell_style(ws.cell(row, col), **header_style)
            row += 1
        for data in datas:
            for col, header in enumerate(headers, start_col):
                ws.cell(row, col).value = data.get(header)
                if self.basic_style or content_style: self._apply_cell_style(ws.cell(row, col), **content_style)
            row += 1
    
    def write_cols(self, sheet: str="", position: tuple = (1, 1), datas: dict={}, headers: list=[], show_headers: bool = True, header_style: dict = {}, content_style: dict = {}):
        '''
        Write table data to sheet.
        position: tuple of (start_row, start_col)
        datas: dict of (header, list) pairs, where each list is a column of data
        headers: list of column headers
        header_style: style dict for header cells (overrides basic_style)
        content_style: style dict for content cells (overrides basic_style)
        '''
        ws = self.get_sheet(sheet)
        start_row, start_col = position

        # for key, value in enumerate(list, int) -> (int, list[0]), (int+1, list[1]), (int+2, list[2]), ...
        data_start_row = start_row + 1 if show_headers else start_row
        for col, header in enumerate(headers, start_col):
            if show_headers:
                ws.cell(start_row, col).value = header
                if self.basic_style or header_style: self._apply_cell_style(ws.cell(start_row, col), **header_style)
            for row, value in enumerate(datas.get(header, []), data_start_row):
                ws.cell(row, col).value = value
                if self.basic_style or content_style: self._apply_cell_style(ws.cell(row, col), **content_style)
    
    def write_cell(self, sheet: str="", position: tuple=(1,1), content: any=None, **kwargs):
        '''
        Write content to cell.
        position: tuple of (row, col)
        content: content to write to cell
        kwargs: keyword arguments for cell style
        '''
        ws = self.get_sheet(sheet)
        row, col = position
        ws.cell(row, col).value = content
        if self.basic_style or kwargs:
            self._apply_cell_style(ws.cell(row, col), **kwargs)
    
    def _apply_cell_style(self, cell, **kwargs):
        merged_kwargs = {**self.basic_style, **kwargs}
        kwargs = merged_kwargs
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
    
    # Styling Examples
    ws = excel_file.get_sheet("styling_examples")
    row = 1
    
    # Font examples
    excel_file.write_cell("styling_examples", (row, 1), "Font Examples", bold=True, size=14)
    row += 2
    
    excel_file.write_cell("styling_examples", (row, 1), "Bold", bold=True)
    excel_file.write_cell("styling_examples", (row, 2), "Italic", italic=True)
    excel_file.write_cell("styling_examples", (row, 3), "Underline", underline="single")
    excel_file.write_cell("styling_examples", (row, 4), "Large Size", size=16)
    excel_file.write_cell("styling_examples", (row, 5), "Arial Font", name="Arial")
    excel_file.write_cell("styling_examples", (row, 6), "Red Color", color="FF0000")
    excel_file.write_cell("styling_examples", (row, 7), "Bold+Italic+Blue", bold=True, italic=True, color="0000FF")
    row += 2
    
    # Fill examples
    excel_file.write_cell("styling_examples", (row, 1), "Fill Examples", bold=True, size=14)
    row += 2
    
    excel_file.write_cell("styling_examples", (row, 1), "Yellow Fill", fgColor="FFFF00")
    excel_file.write_cell("styling_examples", (row, 2), "Green Fill", fgColor="00FF00")
    excel_file.write_cell("styling_examples", (row, 3), "Blue Fill", fgColor="0000FF")
    excel_file.write_cell("styling_examples", (row, 4), "Red Fill", fgColor="FF0000")
    excel_file.write_cell("styling_examples", (row, 5), "Pattern Fill", fill_type="lightGray", fgColor="CCCCCC")
    row += 2
    
    # Alignment examples
    excel_file.write_cell("styling_examples", (row, 1), "Alignment Examples", bold=True, size=14)
    row += 2
    
    excel_file.write_cell("styling_examples", (row, 1), "Left", horizontal="left")
    excel_file.write_cell("styling_examples", (row, 2), "Center", horizontal="center")
    excel_file.write_cell("styling_examples", (row, 3), "Right", horizontal="right")
    excel_file.write_cell("styling_examples", (row, 4), "Top", vertical="top")
    excel_file.write_cell("styling_examples", (row, 5), "Middle", vertical="center")
    excel_file.write_cell("styling_examples", (row, 6), "Bottom", vertical="bottom")
    excel_file.write_cell("styling_examples", (row, 7), "Wrap Text", wrapText=True, horizontal="left")
    ws.cell(row, 7).value = "This is a long text that will wrap to multiple lines"
    row += 2
    
    # Border examples
    excel_file.write_cell("styling_examples", (row, 1), "Border Examples", bold=True, size=14)
    row += 2
    
    excel_file.write_cell("styling_examples", (row, 1), "Thin Border", border_style="thin")
    excel_file.write_cell("styling_examples", (row, 2), "Thick Border", border_style="thick")
    excel_file.write_cell("styling_examples", (row, 3), "Dashed Border", border_style="dashed")
    excel_file.write_cell("styling_examples", (row, 4), "Red Border", border_style="thin", border_color="FF0000")
    excel_file.write_cell("styling_examples", (row, 5), "Double Border", border_style="double")
    row += 2
    
    # Number format examples
    excel_file.write_cell("styling_examples", (row, 1), "Number Format Examples", bold=True, size=14)
    row += 2
    
    excel_file.write_cell("styling_examples", (row, 1), 1234.56, number_format="#,##0.00")
    excel_file.write_cell("styling_examples", (row, 2), 0.15, number_format="0.00%")
    excel_file.write_cell("styling_examples", (row, 3), 1234.56, number_format="$#,##0.00")
    excel_file.write_cell("styling_examples", (row, 4), "2024-01-15", number_format="yyyy-mm-dd")
    row += 2
    
    # Combined examples
    excel_file.write_cell("styling_examples", (row, 1), "Combined Examples", bold=True, size=14)
    row += 2
    
    excel_file.write_cell("styling_examples", (row, 1), "Bold+Yellow+Center", bold=True, fgColor="FFFF00", horizontal="center")
    excel_file.write_cell("styling_examples", (row, 2), "Italic+Blue+Right", italic=True, color="0000FF", horizontal="right")
    excel_file.write_cell("styling_examples", (row, 3), "Bold+Green+Border", bold=True, fgColor="00FF00", border_style="thin")
    excel_file.write_cell("styling_examples", (row, 4), "Large+Red+Center", size=16, color="FF0000", horizontal="center", vertical="center")
    
    excel_file.save()
    print("Styling examples written to test.xlsx")