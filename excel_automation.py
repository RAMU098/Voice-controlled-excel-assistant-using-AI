"""
Excel Automation Module
Handles all Excel operations: data entry, formatting, formulas, charts, filtering, sorting
"""
import os
import re
from datetime import datetime
from typing import Optional, Dict, Any, List, Tuple
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Fill, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, AreaChart, Reference
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from config import Config


class ExcelAutomation:
    """
    Excel Automation class that handles all spreadsheet operations.
    """
    
    def __init__(self, file_path: Optional[str] = None):
        """Initialize Excel automation with optional file path."""
        self.file_path = file_path
        self.workbook = None
        self.active_sheet = None
        
        if file_path and os.path.exists(file_path):
            self.load_workbook(file_path)
        else:
            self.create_new_workbook()
    
    def create_new_workbook(self) -> str:
        """Create a new Excel workbook."""
        self.workbook = Workbook()
        self.active_sheet = self.workbook.active
        self.active_sheet.title = "Sheet1"
        
        # Generate unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.file_path = os.path.join(Config.OUTPUT_FOLDER, f"workbook_{timestamp}.xlsx")
        
        return f"Created new workbook: {os.path.basename(self.file_path)}"
    
    def load_workbook(self, file_path: str) -> str:
        """Load an existing Excel workbook."""
        try:
            self.workbook = load_workbook(file_path)
            self.active_sheet = self.workbook.active
            self.file_path = file_path
            return f"Loaded workbook: {os.path.basename(file_path)}"
        except Exception as e:
            return f"Error loading workbook: {str(e)}"
    
    def save_workbook(self, file_path: Optional[str] = None) -> str:
        """Save the current workbook."""
        try:
            save_path = file_path or self.file_path
            if not save_path:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_path = os.path.join(Config.OUTPUT_FOLDER, f"workbook_{timestamp}.xlsx")
            
            self.workbook.save(save_path)
            self.file_path = save_path
            return f"Saved workbook to: {os.path.basename(save_path)}"
        except Exception as e:
            return f"Error saving workbook: {str(e)}"
    
    def get_file_path(self) -> str:
        """Get the current file path."""
        return self.file_path
    
    # ===================== SHEET OPERATIONS =====================
    
    def create_sheet(self, name: str) -> str:
        """Create a new sheet in the workbook."""
        try:
            if name in self.workbook.sheetnames:
                return f"Sheet '{name}' already exists"
            
            self.workbook.create_sheet(name)
            self.active_sheet = self.workbook[name]
            return f"Created new sheet: {name}"
        except Exception as e:
            return f"Error creating sheet: {str(e)}"
    
    def delete_sheet(self, name: str) -> str:
        """Delete a sheet from the workbook."""
        try:
            if name not in self.workbook.sheetnames:
                return f"Sheet '{name}' not found"
            
            del self.workbook[name]
            self.active_sheet = self.workbook.active
            return f"Deleted sheet: {name}"
        except Exception as e:
            return f"Error deleting sheet: {str(e)}"
    
    def rename_sheet(self, old_name: str, new_name: str) -> str:
        """Rename a sheet."""
        try:
            if old_name not in self.workbook.sheetnames:
                return f"Sheet '{old_name}' not found"
            
            self.workbook[old_name].title = new_name
            return f"Renamed sheet '{old_name}' to '{new_name}'"
        except Exception as e:
            return f"Error renaming sheet: {str(e)}"
    
    def switch_sheet(self, name: str) -> str:
        """Switch to a different sheet (case-insensitive)."""
        try:
            # Case-insensitive sheet name matching
            name_lower = name.lower()
            matched_sheet = None
            for sheet_name in self.workbook.sheetnames:
                if sheet_name.lower() == name_lower:
                    matched_sheet = sheet_name
                    break
            
            if not matched_sheet:
                return f"Sheet '{name}' not found. Available sheets: {', '.join(self.workbook.sheetnames)}"
            
            self.active_sheet = self.workbook[matched_sheet]
            return f"Switched to sheet: {matched_sheet}"
        except Exception as e:
            return f"Error switching sheet: {str(e)}"
    
    def list_sheets(self) -> str:
        """List all sheets in the workbook."""
        sheets = self.workbook.sheetnames
        return f"Available sheets: {', '.join(sheets)}"
    
    # ===================== DATA ENTRY OPERATIONS =====================
    
    def write_cell(self, cell: str, value: Any) -> str:
        """Write a value to a specific cell."""
        try:
            self.active_sheet[cell] = value
            return f"Written '{value}' to cell {cell}"
        except Exception as e:
            return f"Error writing to cell: {str(e)}"
    
    def write_range(self, start_cell: str, data: List[List[Any]]) -> str:
        """Write data to a range of cells starting from start_cell."""
        try:
            start_row = int(re.findall(r'\d+', start_cell)[0])
            start_col = column_index_from_string(re.findall(r'[A-Za-z]+', start_cell)[0])
            
            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    cell = self.active_sheet.cell(row=start_row + row_idx, column=start_col + col_idx)
                    cell.value = value
            
            return f"Written data to range starting at {start_cell}"
        except Exception as e:
            return f"Error writing to range: {str(e)}"
    
    def read_cell(self, cell: str) -> str:
        """Read value from a specific cell."""
        try:
            value = self.active_sheet[cell].value
            return f"Cell {cell} contains: {value}"
        except Exception as e:
            return f"Error reading cell: {str(e)}"
    
    def read_range(self, start_cell: str, end_cell: str) -> str:
        """Read values from a range of cells."""
        try:
            values = []
            for row in self.active_sheet[start_cell:end_cell]:
                row_values = [cell.value for cell in row]
                values.append(row_values)
            return f"Range {start_cell}:{end_cell} contains: {values}"
        except Exception as e:
            return f"Error reading range: {str(e)}"
    
    def clear_cell(self, cell: str) -> str:
        """Clear a specific cell."""
        try:
            self.active_sheet[cell].value = None
            return f"Cleared cell {cell}"
        except Exception as e:
            return f"Error clearing cell: {str(e)}"
    
    def clear_range(self, start_cell: str, end_cell: str) -> str:
        """Clear a range of cells."""
        try:
            for row in self.active_sheet[start_cell:end_cell]:
                for cell in row:
                    cell.value = None
            return f"Cleared range {start_cell}:{end_cell}"
        except Exception as e:
            return f"Error clearing range: {str(e)}"
    
    # ===================== ROW/COLUMN OPERATIONS =====================
    
    def insert_row(self, row_number: int, count: int = 1) -> str:
        """Insert row(s) at specified position."""
        try:
            self.active_sheet.insert_rows(row_number, count)
            return f"Inserted {count} row(s) at row {row_number}"
        except Exception as e:
            return f"Error inserting row: {str(e)}"
    
    def delete_row(self, row_number: int, count: int = 1) -> str:
        """Delete row(s) at specified position."""
        try:
            self.active_sheet.delete_rows(row_number, count)
            return f"Deleted {count} row(s) at row {row_number}"
        except Exception as e:
            return f"Error deleting row: {str(e)}"
    
    def insert_column(self, column: str, count: int = 1) -> str:
        """Insert column(s) at specified position."""
        try:
            col_idx = column_index_from_string(column.upper())
            self.active_sheet.insert_cols(col_idx, count)
            return f"Inserted {count} column(s) at column {column}"
        except Exception as e:
            return f"Error inserting column: {str(e)}"
    
    def delete_column(self, column: str, count: int = 1) -> str:
        """Delete column(s) at specified position."""
        try:
            col_idx = column_index_from_string(column.upper())
            self.active_sheet.delete_cols(col_idx, count)
            return f"Deleted {count} column(s) at column {column}"
        except Exception as e:
            return f"Error deleting column: {str(e)}"
    
    def set_column_width(self, column: str, width: float) -> str:
        """Set column width."""
        try:
            self.active_sheet.column_dimensions[column.upper()].width = width
            return f"Set column {column} width to {width}"
        except Exception as e:
            return f"Error setting column width: {str(e)}"
    
    def set_row_height(self, row_number: int, height: float) -> str:
        """Set row height."""
        try:
            self.active_sheet.row_dimensions[row_number].height = height
            return f"Set row {row_number} height to {height}"
        except Exception as e:
            return f"Error setting row height: {str(e)}"
    
    # ===================== FORMATTING OPERATIONS =====================
    
    def format_cell(self, cell: str, **kwargs) -> str:
        """
        Format a cell with various options.
        kwargs: bold, italic, underline, font_size, font_color, bg_color, align
        """
        try:
            target_cell = self.active_sheet[cell]
            
            # Font formatting
            font_props = {}
            if 'bold' in kwargs:
                font_props['bold'] = kwargs['bold']
            if 'italic' in kwargs:
                font_props['italic'] = kwargs['italic']
            if 'underline' in kwargs:
                font_props['underline'] = 'single' if kwargs['underline'] else None
            if 'font_size' in kwargs:
                font_props['size'] = kwargs['font_size']
            if 'font_color' in kwargs:
                font_props['color'] = kwargs['font_color']
            if 'font_name' in kwargs:
                font_props['name'] = kwargs['font_name']
            
            if font_props:
                target_cell.font = Font(**font_props)
            
            # Background color
            if 'bg_color' in kwargs:
                target_cell.fill = PatternFill(start_color=kwargs['bg_color'],
                                                end_color=kwargs['bg_color'],
                                                fill_type='solid')
            
            # Alignment
            if 'align' in kwargs:
                alignment = kwargs['align'].lower()
                if alignment == 'center':
                    target_cell.alignment = Alignment(horizontal='center', vertical='center')
                elif alignment == 'left':
                    target_cell.alignment = Alignment(horizontal='left')
                elif alignment == 'right':
                    target_cell.alignment = Alignment(horizontal='right')
            
            return f"Formatted cell {cell}"
        except Exception as e:
            return f"Error formatting cell: {str(e)}"
    
    def format_range(self, start_cell: str, end_cell: str, **kwargs) -> str:
        """Format a range of cells."""
        try:
            for row in self.active_sheet[start_cell:end_cell]:
                for cell in row:
                    self.format_cell(cell.coordinate, **kwargs)
            return f"Formatted range {start_cell}:{end_cell}"
        except Exception as e:
            return f"Error formatting range: {str(e)}"
    
    def format_as_currency(self, cell: str, currency: str = "$") -> str:
        """Format cell as currency."""
        try:
            target_cell = self.active_sheet[cell]
            if currency == "$":
                target_cell.number_format = '$#,##0.00'
            elif currency == "€":
                target_cell.number_format = '€#,##0.00'
            elif currency == "₹":
                target_cell.number_format = '₹#,##0.00'
            else:
                target_cell.number_format = f'{currency}#,##0.00'
            return f"Formatted cell {cell} as currency ({currency})"
        except Exception as e:
            return f"Error formatting as currency: {str(e)}"
    
    def format_as_percentage(self, cell: str) -> str:
        """Format cell as percentage."""
        try:
            self.active_sheet[cell].number_format = '0.00%'
            return f"Formatted cell {cell} as percentage"
        except Exception as e:
            return f"Error formatting as percentage: {str(e)}"
    
    def format_as_date(self, cell: str) -> str:
        """Format cell as date."""
        try:
            self.active_sheet[cell].number_format = 'YYYY-MM-DD'
            return f"Formatted cell {cell} as date"
        except Exception as e:
            return f"Error formatting as date: {str(e)}"
    
    def add_border(self, cell: str, border_style: str = 'thin') -> str:
        """Add border to a cell."""
        try:
            side = Side(style=border_style)
            border = Border(left=side, right=side, top=side, bottom=side)
            self.active_sheet[cell].border = border
            return f"Added {border_style} border to cell {cell}"
        except Exception as e:
            return f"Error adding border: {str(e)}"
    
    def add_border_range(self, start_cell: str, end_cell: str, border_style: str = 'thin') -> str:
        """Add border to a range of cells."""
        try:
            side = Side(style=border_style)
            border = Border(left=side, right=side, top=side, bottom=side)
            for row in self.active_sheet[start_cell:end_cell]:
                for cell in row:
                    cell.border = border
            return f"Added {border_style} border to range {start_cell}:{end_cell}"
        except Exception as e:
            return f"Error adding border to range: {str(e)}"
    
    def merge_cells(self, start_cell: str, end_cell: str) -> str:
        """Merge a range of cells."""
        try:
            self.active_sheet.merge_cells(f"{start_cell}:{end_cell}")
            return f"Merged cells {start_cell}:{end_cell}"
        except Exception as e:
            return f"Error merging cells: {str(e)}"
    
    def unmerge_cells(self, start_cell: str, end_cell: str) -> str:
        """Unmerge a range of cells."""
        try:
            self.active_sheet.unmerge_cells(f"{start_cell}:{end_cell}")
            return f"Unmerged cells {start_cell}:{end_cell}"
        except Exception as e:
            return f"Error unmerging cells: {str(e)}"
    
    # ===================== FORMULA OPERATIONS =====================
    
    def insert_formula(self, cell: str, formula: str) -> str:
        """Insert a formula into a cell."""
        try:
            # Ensure formula starts with =
            if not formula.startswith('='):
                formula = '=' + formula
            self.active_sheet[cell] = formula
            return f"Inserted formula '{formula}' in cell {cell}"
        except Exception as e:
            return f"Error inserting formula: {str(e)}"
    
    def sum_range(self, result_cell: str, range_start: str, range_end: str) -> str:
        """Insert SUM formula."""
        try:
            formula = f"=SUM({range_start}:{range_end})"
            self.active_sheet[result_cell] = formula
            return f"Added SUM formula in {result_cell} for range {range_start}:{range_end}"
        except Exception as e:
            return f"Error adding SUM: {str(e)}"
    
    def average_range(self, result_cell: str, range_start: str, range_end: str) -> str:
        """Insert AVERAGE formula."""
        try:
            formula = f"=AVERAGE({range_start}:{range_end})"
            self.active_sheet[result_cell] = formula
            return f"Added AVERAGE formula in {result_cell} for range {range_start}:{range_end}"
        except Exception as e:
            return f"Error adding AVERAGE: {str(e)}"
    
    def count_range(self, result_cell: str, range_start: str, range_end: str) -> str:
        """Insert COUNT formula."""
        try:
            formula = f"=COUNT({range_start}:{range_end})"
            self.active_sheet[result_cell] = formula
            return f"Added COUNT formula in {result_cell} for range {range_start}:{range_end}"
        except Exception as e:
            return f"Error adding COUNT: {str(e)}"
    
    def max_range(self, result_cell: str, range_start: str, range_end: str) -> str:
        """Insert MAX formula."""
        try:
            formula = f"=MAX({range_start}:{range_end})"
            self.active_sheet[result_cell] = formula
            return f"Added MAX formula in {result_cell} for range {range_start}:{range_end}"
        except Exception as e:
            return f"Error adding MAX: {str(e)}"
    
    def min_range(self, result_cell: str, range_start: str, range_end: str) -> str:
        """Insert MIN formula."""
        try:
            formula = f"=MIN({range_start}:{range_end})"
            self.active_sheet[result_cell] = formula
            return f"Added MIN formula in {result_cell} for range {range_start}:{range_end}"
        except Exception as e:
            return f"Error adding MIN: {str(e)}"
    
    def sum_row(self, row_number: int) -> str:
        """Sum all numeric values in a row."""
        try:
            # Find the extent of data in the row
            max_col = self.active_sheet.max_column or 1
            
            # Find first and last columns with data
            first_col = None
            last_col = None
            
            for col in range(1, max_col + 1):
                cell_value = self.active_sheet.cell(row=row_number, column=col).value
                if cell_value is not None:
                    if first_col is None:
                        first_col = col
                    last_col = col
            
            if first_col is None:
                return f"Row {row_number} is empty"
            
            # Create range
            start_col = get_column_letter(first_col)
            end_col = get_column_letter(last_col)
            
            # Put result in the next column after data
            result_col = get_column_letter(last_col + 1)
            result_cell = f"{result_col}{row_number}"
            
            formula = f"=SUM({start_col}{row_number}:{end_col}{row_number})"
            self.active_sheet[result_cell] = formula
            
            return f"Added SUM formula for row {row_number} in cell {result_cell} (summing {start_col}{row_number}:{end_col}{row_number})"
        except Exception as e:
            return f"Error summing row: {str(e)}"
    
    def sum_column(self, column: str) -> str:
        """Sum all numeric values in a column."""
        try:
            col_idx = column_index_from_string(column.upper())
            max_row = self.active_sheet.max_row or 1
            
            # Find first and last rows with data in this column
            first_row = None
            last_row = None
            
            for row in range(1, max_row + 1):
                cell_value = self.active_sheet.cell(row=row, column=col_idx).value
                if cell_value is not None:
                    if first_row is None:
                        first_row = row
                    last_row = row
            
            if first_row is None:
                return f"Column {column} is empty"
            
            # Put result in the next row after data
            result_row = last_row + 1
            result_cell = f"{column.upper()}{result_row}"
            
            formula = f"=SUM({column.upper()}{first_row}:{column.upper()}{last_row})"
            self.active_sheet[result_cell] = formula
            
            return f"Added SUM formula for column {column.upper()} in cell {result_cell} (summing {column.upper()}{first_row}:{column.upper()}{last_row})"
        except Exception as e:
            return f"Error summing column: {str(e)}"
    
    def average_column(self, column: str) -> str:
        """Average all numeric values in a column."""
        try:
            col_idx = column_index_from_string(column.upper())
            max_row = self.active_sheet.max_row or 1
            
            first_row = None
            last_row = None
            
            for row in range(1, max_row + 1):
                cell_value = self.active_sheet.cell(row=row, column=col_idx).value
                if cell_value is not None:
                    if first_row is None:
                        first_row = row
                    last_row = row
            
            if first_row is None:
                return f"Column {column} is empty"
            
            result_row = last_row + 1
            result_cell = f"{column.upper()}{result_row}"
            
            formula = f"=AVERAGE({column.upper()}{first_row}:{column.upper()}{last_row})"
            self.active_sheet[result_cell] = formula
            
            return f"Added AVERAGE formula for column {column.upper()} in cell {result_cell}"
        except Exception as e:
            return f"Error averaging column: {str(e)}"
    
    def format_column(self, column: str, **kwargs) -> str:
        """Format an entire column."""
        try:
            col_idx = column_index_from_string(column.upper())
            max_row = self.active_sheet.max_row or 1
            
            formatted_count = 0
            for row in range(1, max_row + 1):
                cell = self.active_sheet.cell(row=row, column=col_idx)
                if cell.value is not None:
                    cell_ref = f"{column.upper()}{row}"
                    self.format_cell(cell_ref, **kwargs)
                    formatted_count += 1
            
            return f"Formatted {formatted_count} cells in column {column.upper()}"
        except Exception as e:
            return f"Error formatting column: {str(e)}"
    
    def average_row(self, row_number: int) -> str:
        """Average all numeric values in a row."""
        try:
            max_col = self.active_sheet.max_column or 1
            
            first_col = None
            last_col = None
            
            for col in range(1, max_col + 1):
                cell_value = self.active_sheet.cell(row=row_number, column=col).value
                if cell_value is not None:
                    if first_col is None:
                        first_col = col
                    last_col = col
            
            if first_col is None:
                return f"Row {row_number} is empty"
            
            start_col = get_column_letter(first_col)
            end_col = get_column_letter(last_col)
            result_col = get_column_letter(last_col + 1)
            result_cell = f"{result_col}{row_number}"
            
            formula = f"=AVERAGE({start_col}{row_number}:{end_col}{row_number})"
            self.active_sheet[result_cell] = formula
            
            return f"Added AVERAGE formula for row {row_number} in cell {result_cell}"
        except Exception as e:
            return f"Error averaging row: {str(e)}"
    
    def if_formula(self, cell: str, condition: str, true_value: str, false_value: str) -> str:
        """Insert IF formula."""
        try:
            formula = f"=IF({condition},{true_value},{false_value})"
            self.active_sheet[cell] = formula
            return f"Added IF formula in cell {cell}"
        except Exception as e:
            return f"Error adding IF formula: {str(e)}"
    
    def vlookup_formula(self, cell: str, lookup_value: str, table_range: str, 
                        col_index: int, exact_match: bool = True) -> str:
        """Insert VLOOKUP formula."""
        try:
            match_type = "FALSE" if exact_match else "TRUE"
            formula = f"=VLOOKUP({lookup_value},{table_range},{col_index},{match_type})"
            self.active_sheet[cell] = formula
            return f"Added VLOOKUP formula in cell {cell}"
        except Exception as e:
            return f"Error adding VLOOKUP: {str(e)}"
    
    # ===================== CHART OPERATIONS =====================
    
    def create_bar_chart(self, data_range: str, title: str = "Bar Chart", 
                         position: str = "E5") -> str:
        """Create a bar chart."""
        try:
            # Parse range
            parts = data_range.split(':')
            start_cell = parts[0]
            end_cell = parts[1] if len(parts) > 1 else start_cell
            
            start_col = column_index_from_string(re.findall(r'[A-Za-z]+', start_cell)[0])
            start_row = int(re.findall(r'\d+', start_cell)[0])
            end_col = column_index_from_string(re.findall(r'[A-Za-z]+', end_cell)[0])
            end_row = int(re.findall(r'\d+', end_cell)[0])
            
            chart = BarChart()
            chart.title = title
            chart.type = "col"
            
            data = Reference(self.active_sheet, min_col=start_col + 1, min_row=start_row,
                           max_col=end_col, max_row=end_row)
            categories = Reference(self.active_sheet, min_col=start_col, min_row=start_row + 1,
                                  max_row=end_row)
            
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            
            self.active_sheet.add_chart(chart, position)
            return f"Created bar chart '{title}' at position {position}"
        except Exception as e:
            return f"Error creating bar chart: {str(e)}"
    
    def create_pie_chart(self, data_range: str, title: str = "Pie Chart",
                         position: str = "E5") -> str:
        """Create a pie chart."""
        try:
            parts = data_range.split(':')
            start_cell = parts[0]
            end_cell = parts[1] if len(parts) > 1 else start_cell
            
            start_col = column_index_from_string(re.findall(r'[A-Za-z]+', start_cell)[0])
            start_row = int(re.findall(r'\d+', start_cell)[0])
            end_col = column_index_from_string(re.findall(r'[A-Za-z]+', end_cell)[0])
            end_row = int(re.findall(r'\d+', end_cell)[0])
            
            chart = PieChart()
            chart.title = title
            
            data = Reference(self.active_sheet, min_col=start_col + 1, min_row=start_row,
                           max_row=end_row)
            labels = Reference(self.active_sheet, min_col=start_col, min_row=start_row + 1,
                             max_row=end_row)
            
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
            
            self.active_sheet.add_chart(chart, position)
            return f"Created pie chart '{title}' at position {position}"
        except Exception as e:
            return f"Error creating pie chart: {str(e)}"
    
    def create_line_chart(self, data_range: str, title: str = "Line Chart",
                          position: str = "E5") -> str:
        """Create a line chart."""
        try:
            parts = data_range.split(':')
            start_cell = parts[0]
            end_cell = parts[1] if len(parts) > 1 else start_cell
            
            start_col = column_index_from_string(re.findall(r'[A-Za-z]+', start_cell)[0])
            start_row = int(re.findall(r'\d+', start_cell)[0])
            end_col = column_index_from_string(re.findall(r'[A-Za-z]+', end_cell)[0])
            end_row = int(re.findall(r'\d+', end_cell)[0])
            
            chart = LineChart()
            chart.title = title
            chart.style = 10
            
            data = Reference(self.active_sheet, min_col=start_col + 1, min_row=start_row,
                           max_col=end_col, max_row=end_row)
            categories = Reference(self.active_sheet, min_col=start_col, min_row=start_row + 1,
                                  max_row=end_row)
            
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            
            self.active_sheet.add_chart(chart, position)
            return f"Created line chart '{title}' at position {position}"
        except Exception as e:
            return f"Error creating line chart: {str(e)}"
    
    # ===================== SORTING & FILTERING =====================
    
    def sort_column(self, column: str, ascending: bool = True) -> str:
        """Sort data by a column (using pandas for efficiency)."""
        try:
            # Read data into pandas
            data = []
            max_col = self.active_sheet.max_column
            max_row = self.active_sheet.max_row
            
            if max_row is None or max_row < 2:
                return "No data to sort"
            
            for row in self.active_sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True):
                # Filter out completely empty rows
                if any(cell is not None for cell in row):
                    data.append(list(row))
            
            if len(data) < 2:
                return "Not enough data to sort"
            
            # Get column index
            col_idx = column_index_from_string(column.upper()) - 1
            
            if col_idx >= len(data[0]):
                return f"Column {column} is out of range"
            
            # Create DataFrame with string headers to avoid issues
            headers = [str(h) if h is not None else f"Col{i}" for i, h in enumerate(data[0])]
            df = pd.DataFrame(data[1:], columns=headers)
            col_name = headers[col_idx]
            
            # Convert column to numeric if possible for proper sorting
            try:
                df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
            except:
                pass
            
            # Sort
            df = df.sort_values(by=col_name, ascending=ascending, na_position='last')
            
            # Clear existing data
            for row in self.active_sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
                for cell in row:
                    cell.value = None
            
            # Write back sorted data
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    self.active_sheet.cell(row=r_idx, column=c_idx, value=value)
            
            order = "ascending" if ascending else "descending"
            return f"Sorted data by column {column} in {order} order"
        except Exception as e:
            return f"Error sorting: {str(e)}"
    
    def filter_data(self, column: str, value: Any) -> str:
        """Filter data by a column value."""
        try:
            # Read data into pandas
            data = []
            for row in self.active_sheet.iter_rows(values_only=True):
                data.append(row)
            
            if not data:
                return "No data to filter"
            
            # Create DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])
            col_idx = column_index_from_string(column.upper()) - 1
            col_name = df.columns[col_idx]
            
            # Filter
            df_filtered = df[df[col_name] == value]
            
            # Clear existing data
            for row in self.active_sheet.iter_rows():
                for cell in row:
                    cell.value = None
            
            # Write back filtered data
            for r_idx, row in enumerate(dataframe_to_rows(df_filtered, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    self.active_sheet.cell(row=r_idx, column=c_idx, value=value)
            
            return f"Filtered data where column {column} equals '{value}'"
        except Exception as e:
            return f"Error filtering: {str(e)}"
    
    def auto_filter(self, range_str: str) -> str:
        """Enable auto filter on a range."""
        try:
            self.active_sheet.auto_filter.ref = range_str
            return f"Enabled auto filter on range {range_str}"
        except Exception as e:
            return f"Error enabling auto filter: {str(e)}"
    
    # ===================== DATA IMPORT/EXPORT =====================
    
    def import_csv(self, csv_path: str, start_cell: str = "A1") -> str:
        """Import data from CSV file."""
        try:
            df = pd.read_csv(csv_path)
            
            start_row = int(re.findall(r'\d+', start_cell)[0])
            start_col = column_index_from_string(re.findall(r'[A-Za-z]+', start_cell)[0])
            
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
                for c_idx, value in enumerate(row, 1):
                    self.active_sheet.cell(row=start_row + r_idx, 
                                          column=start_col + c_idx - 1, 
                                          value=value)
            
            return f"Imported CSV data from {os.path.basename(csv_path)}"
        except Exception as e:
            return f"Error importing CSV: {str(e)}"
    
    def export_csv(self, csv_path: str) -> str:
        """Export sheet data to CSV."""
        try:
            data = []
            for row in self.active_sheet.iter_rows(values_only=True):
                data.append(row)
            
            df = pd.DataFrame(data[1:], columns=data[0])
            df.to_csv(csv_path, index=False)
            
            return f"Exported data to {os.path.basename(csv_path)}"
        except Exception as e:
            return f"Error exporting CSV: {str(e)}"
    
    # ===================== SAMPLE DATA =====================
    
    def create_sample_data(self, data_type: str = "sales") -> str:
        """Create sample data for demonstration."""
        try:
            if data_type == "sales":
                headers = ["Product", "Q1", "Q2", "Q3", "Q4", "Total"]
                data = [
                    ["Laptop", 15000, 18000, 22000, 25000],
                    ["Phone", 12000, 14000, 16000, 18000],
                    ["Tablet", 8000, 9500, 11000, 13000],
                    ["Monitor", 5000, 6000, 7500, 8500],
                    ["Keyboard", 2000, 2500, 3000, 3500],
                ]
            elif data_type == "employees":
                headers = ["Name", "Department", "Salary", "Experience", "Rating"]
                data = [
                    ["John Smith", "Engineering", 75000, 5, 4.5],
                    ["Jane Doe", "Marketing", 65000, 3, 4.2],
                    ["Bob Wilson", "Sales", 70000, 7, 4.8],
                    ["Alice Brown", "HR", 55000, 2, 4.0],
                    ["Charlie Davis", "Engineering", 85000, 8, 4.7],
                ]
            elif data_type == "inventory":
                headers = ["Item", "Category", "Quantity", "Price", "Value"]
                data = [
                    ["Widget A", "Electronics", 150, 25.99],
                    ["Widget B", "Electronics", 200, 15.50],
                    ["Gadget X", "Accessories", 75, 45.00],
                    ["Gadget Y", "Accessories", 120, 35.75],
                    ["Tool Z", "Hardware", 50, 89.99],
                ]
            else:
                headers = ["Column A", "Column B", "Column C"]
                data = [
                    [100, 200, 300],
                    [150, 250, 350],
                    [200, 300, 400],
                ]
            
            # Write headers
            for col, header in enumerate(headers, 1):
                cell = self.active_sheet.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF")
            
            # Write data
            for row_idx, row_data in enumerate(data, 2):
                for col_idx, value in enumerate(row_data, 1):
                    self.active_sheet.cell(row=row_idx, column=col_idx, value=value)
            
            # Add formulas for totals if sales data
            if data_type == "sales":
                for row_idx in range(2, len(data) + 2):
                    formula = f"=SUM(B{row_idx}:E{row_idx})"
                    self.active_sheet.cell(row=row_idx, column=6, value=formula)
            
            # Auto-adjust column widths
            for col in range(1, len(headers) + 1):
                self.active_sheet.column_dimensions[get_column_letter(col)].width = 15
            
            return f"Created sample {data_type} data with {len(data)} rows"
        except Exception as e:
            return f"Error creating sample data: {str(e)}"
    
    # ===================== UTILITY METHODS =====================
    
    def get_sheet_info(self) -> Dict[str, Any]:
        """Get information about the current sheet."""
        try:
            max_row = self.active_sheet.max_row
            max_col = self.active_sheet.max_column
            
            return {
                "sheet_name": self.active_sheet.title,
                "max_row": max_row,
                "max_column": max_col,
                "max_column_letter": get_column_letter(max_col) if max_col else "A",
                "all_sheets": self.workbook.sheetnames
            }
        except Exception as e:
            return {"error": str(e)}
    
    def get_data_preview(self, rows: int = 10) -> List[List[Any]]:
        """Get a preview of the sheet data with calculated formula values."""
        try:
            data = []
            max_col = self.active_sheet.max_column or 1
            
            for row_idx in range(1, min(rows + 1, (self.active_sheet.max_row or 0) + 1)):
                row_data = []
                for col_idx in range(1, max_col + 1):
                    cell = self.active_sheet.cell(row=row_idx, column=col_idx)
                    value = cell.value
                    
                    # If it's a formula, calculate the value
                    if isinstance(value, str) and value.startswith('='):
                        calculated = self._calculate_formula(value)
                        row_data.append(calculated)
                    else:
                        row_data.append(value)
                
                if any(v is not None for v in row_data):
                    data.append(row_data)
                elif row_idx <= rows:
                    data.append(row_data)
            
            return data
        except Exception as e:
            return [[f"Error: {str(e)}"]]
    
    def _calculate_formula(self, formula: str) -> Any:
        """Calculate simple Excel formulas (SUM, AVERAGE, MAX, MIN, COUNT)."""
        try:
            formula_upper = formula.upper()
            
            # Extract range from formula like =SUM(B2:E2) or =AVERAGE(A1:A10)
            import re
            range_match = re.search(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', formula_upper)
            
            if not range_match:
                return formula  # Can't parse, return formula string
            
            start_col = range_match.group(1)
            start_row = int(range_match.group(2))
            end_col = range_match.group(3)
            end_row = int(range_match.group(4))
            
            # Collect numeric values from the range
            values = []
            start_col_idx = column_index_from_string(start_col)
            end_col_idx = column_index_from_string(end_col)
            
            for r in range(start_row, end_row + 1):
                for c in range(start_col_idx, end_col_idx + 1):
                    cell_val = self.active_sheet.cell(row=r, column=c).value
                    if isinstance(cell_val, (int, float)):
                        values.append(cell_val)
                    elif isinstance(cell_val, str):
                        # Try to parse as number
                        try:
                            values.append(float(cell_val))
                        except ValueError:
                            pass  # Skip non-numeric strings
            
            if not values:
                return 0
            
            # Calculate based on formula type
            if 'SUM' in formula_upper:
                result = sum(values)
            elif 'AVERAGE' in formula_upper or 'AVG' in formula_upper:
                result = sum(values) / len(values)
            elif 'MAX' in formula_upper:
                result = max(values)
            elif 'MIN' in formula_upper:
                result = min(values)
            elif 'COUNT' in formula_upper:
                result = len(values)
            else:
                return formula  # Unknown formula type
            
            # Format result nicely
            if isinstance(result, float):
                if result == int(result):
                    return int(result)
                return round(result, 2)
            return result
            
        except Exception as e:
            return formula  # Return original formula on error
    
    # ===================== ADDITIONAL OPERATIONS =====================
    
    def copy_cell(self, source: str, destination: str) -> str:
        """Copy a cell value to another cell."""
        try:
            source_value = self.active_sheet[source].value
            self.active_sheet[destination] = source_value
            return f"Copied {source} to {destination}"
        except Exception as e:
            return f"Error copying cell: {str(e)}"
    
    def copy_range(self, source_start: str, source_end: str, dest_start: str) -> str:
        """Copy a range of cells to another location."""
        try:
            import re
            # Parse source range
            src_start_col = column_index_from_string(re.findall(r'[A-Za-z]+', source_start)[0])
            src_start_row = int(re.findall(r'\d+', source_start)[0])
            src_end_col = column_index_from_string(re.findall(r'[A-Za-z]+', source_end)[0])
            src_end_row = int(re.findall(r'\d+', source_end)[0])
            
            # Parse destination
            dest_col = column_index_from_string(re.findall(r'[A-Za-z]+', dest_start)[0])
            dest_row = int(re.findall(r'\d+', dest_start)[0])
            
            # Copy values
            for row_offset in range(src_end_row - src_start_row + 1):
                for col_offset in range(src_end_col - src_start_col + 1):
                    src_cell = self.active_sheet.cell(row=src_start_row + row_offset, column=src_start_col + col_offset)
                    dest_cell = self.active_sheet.cell(row=dest_row + row_offset, column=dest_col + col_offset)
                    dest_cell.value = src_cell.value
            
            return f"Copied range {source_start}:{source_end} to {dest_start}"
        except Exception as e:
            return f"Error copying range: {str(e)}"
    
    def find_value(self, search_value: str) -> str:
        """Find a value in the sheet."""
        try:
            found_cells = []
            for row in self.active_sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and str(search_value).lower() in str(cell.value).lower():
                        found_cells.append(cell.coordinate)
            
            if found_cells:
                return f"Found '{search_value}' in cells: {', '.join(found_cells[:10])}" + (f" and {len(found_cells) - 10} more" if len(found_cells) > 10 else "")
            return f"'{search_value}' not found in the sheet"
        except Exception as e:
            return f"Error finding value: {str(e)}"
    
    def replace_value(self, search_value: str, replace_value: str) -> str:
        """Find and replace values in the sheet."""
        try:
            replaced_count = 0
            for row in self.active_sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and str(search_value).lower() in str(cell.value).lower():
                        cell.value = str(cell.value).replace(search_value, replace_value)
                        replaced_count += 1
            
            return f"Replaced {replaced_count} occurrence(s) of '{search_value}' with '{replace_value}'"
        except Exception as e:
            return f"Error replacing value: {str(e)}"
    
    def count_column(self, column: str) -> str:
        """Count non-empty cells in a column."""
        try:
            col_idx = column_index_from_string(column.upper())
            max_row = self.active_sheet.max_row or 1
            
            count = 0
            for row in range(1, max_row + 1):
                if self.active_sheet.cell(row=row, column=col_idx).value is not None:
                    count += 1
            
            return f"Column {column.upper()} has {count} non-empty cells"
        except Exception as e:
            return f"Error counting column: {str(e)}"
    
    def max_column(self, column: str) -> str:
        """Find maximum value in a column."""
        try:
            col_idx = column_index_from_string(column.upper())
            max_row = self.active_sheet.max_row or 1
            
            values = []
            for row in range(1, max_row + 1):
                val = self.active_sheet.cell(row=row, column=col_idx).value
                if isinstance(val, (int, float)):
                    values.append(val)
            
            if values:
                max_val = max(values)
                return f"Maximum value in column {column.upper()} is {max_val}"
            return f"No numeric values in column {column.upper()}"
        except Exception as e:
            return f"Error finding max: {str(e)}"
    
    def min_column(self, column: str) -> str:
        """Find minimum value in a column."""
        try:
            col_idx = column_index_from_string(column.upper())
            max_row = self.active_sheet.max_row or 1
            
            values = []
            for row in range(1, max_row + 1):
                val = self.active_sheet.cell(row=row, column=col_idx).value
                if isinstance(val, (int, float)):
                    values.append(val)
            
            if values:
                min_val = min(values)
                return f"Minimum value in column {column.upper()} is {min_val}"
            return f"No numeric values in column {column.upper()}"
        except Exception as e:
            return f"Error finding min: {str(e)}"
    
    def auto_fit_columns(self) -> str:
        """Auto-fit all column widths based on content."""
        try:
            for column in self.active_sheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50
                self.active_sheet.column_dimensions[column_letter].width = adjusted_width
            
            return "Auto-fitted all column widths"
        except Exception as e:
            return f"Error auto-fitting columns: {str(e)}"
    
    def clear_sheet(self) -> str:
        """Clear all data from the current sheet."""
        try:
            for row in self.active_sheet.iter_rows():
                for cell in row:
                    cell.value = None
            return f"Cleared all data from sheet '{self.active_sheet.title}'"
        except Exception as e:
            return f"Error clearing sheet: {str(e)}"
    
    def freeze_panes(self, cell: str = "A2") -> str:
        """Freeze panes at a specified cell."""
        try:
            self.active_sheet.freeze_panes = cell
            return f"Froze panes at cell {cell}"
        except Exception as e:
            return f"Error freezing panes: {str(e)}"
    
    def unfreeze_panes(self) -> str:
        """Unfreeze all panes."""
        try:
            self.active_sheet.freeze_panes = None
            return "Unfroze all panes"
        except Exception as e:
            return f"Error unfreezing panes: {str(e)}"

