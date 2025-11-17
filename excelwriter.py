import os
import sys
from openpyxl import load_workbook, Workbook

class ExcelFile:
    def __init__(self, filename, basic_style={}):
        self.file_path = self.get_resource_path(filename)
        if os.path.exists(self.file_path):
            self.workbook = load_workbook(self.file_path)
        else:
            self.workbook = Workbook()
            if 'Sheet' in self.workbook.sheetnames:
                self.workbook.remove(self.workbook['Sheet'])

    def get_resource_path(self, relative_path):
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            file_path = os.path.join(exe_dir, relative_path)
            if os.path.exists(file_path):
                return file_path
            try:
                base_path = sys._MEIPASS
            except AttributeError:
                base_path = exe_dir
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)
    
    def save(self):
        self.workbook.save(self.file_path)
    
    def get_sheet(self, sheet):
        if sheet in self.workbook.sheetnames:
            return self.workbook[sheet]
        return self.workbook.create_sheet(title=sheet)
    
    def clear_sheet(self, sheet):
        ws = self.get_sheet(sheet)
        if ws.max_row > 0:
            ws.delete_rows(1, ws.max_row)
    
    def write_rows(self, sheet="", position=(1,1), datas=[], headers=[], show_headers=True, header_style={}, content_style={}):
        ws = self.get_sheet(sheet)
        start_row, start_col = position
        row = start_row
        if show_headers:
            for col, header in enumerate(headers, start_col):
                ws.cell(row, col).value = header
                if header_style:
                    self._apply_cell_style(ws.cell(row, col), **header_style)
            row += 1
        for data in datas:
            for col, header in enumerate(headers, start_col):
                ws.cell(row, col).value = data.get(header)
                if content_style:
                    self._apply_cell_style(ws.cell(row, col), **content_style)
            row += 1
    
    def write_cols(self, sheet="", position=(1, 1), datas={}, headers=[], show_headers=True, header_style={}, content_style={}):
        ws = self.get_sheet(sheet)
        start_row, start_col = position
        data_start_row = start_row + 1 if show_headers else start_row
        for col, header in enumerate(headers, start_col):
            if show_headers:
                ws.cell(start_row, col).value = header
                if header_style:
                    self._apply_cell_style(ws.cell(start_row, col), **header_style)
            for row, value in enumerate(datas.get(header, []), data_start_row):
                ws.cell(row, col).value = value
                if content_style:
                    self._apply_cell_style(ws.cell(row, col), **content_style)
    
    def write_cell(self, sheet="", position=(1,1), content=None, content_style={}):
        ws = self.get_sheet(sheet)
        row, col = position
        ws.cell(row, col).value = content
        if content_style:
            self._apply_cell_style(ws.cell(row, col), **content_style)

    ''' Pattern filling algorithm, but not used for now. Lets solve the real problem first.
    def _apply_cell_style(self, cell, **kwargs):
        font_props = ['name', 'size', 'bold', 'italic', 'underline', 'color']
        font_kwargs = {k: kwargs.pop(k) for k in font_props if k in kwargs}
        if font_kwargs:
            cell.font = Font(**font_kwargs)
        
        fill_props = ['start_color', 'end_color', 'patternType']
        fill_kwargs = {k: kwargs.pop(k) for k in fill_props if k in kwargs}
        if fill_kwargs:
            fill_kwargs.setdefault('patternType', 'solid')
            cell.fill = PatternFill(**fill_kwargs)
        
        align_props = ['horizontal', 'vertical', 'wrap_text']
        align_kwargs = {k: kwargs.pop(k) for k in align_props if k in kwargs}
        if align_kwargs:
            cell.alignment = Alignment(**align_kwargs)
        
        for key, value in kwargs.items():
            if hasattr(cell, key): setattr(cell, key, value)
    '''