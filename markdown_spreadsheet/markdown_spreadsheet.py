import re
from typing import Dict
from sly import Lexer, Parser
from html_generator import HTMLGenerator

class SpreadsheetLexer(Lexer):
    tokens = {'CELL', 'SEPARATOR', 'PLUS', 'TIMES', 'MINUS', 'DIVIDE', 'LPAREN', 'RPAREN', 'NUMBER', 'PIPE','EMPTY_CELL', 'CELL_REFERENCE', 'EQUAL', 'FUNCTION_NAME','SEMICOLON', 'RANGE'}
    # Ignore whitespace characters
    ignore = ' \t'  
    
    EMPTY_CELL = r'(EMPTY)'
    SEPARATOR = r'\|-{3,}\|'    
    EQUAL = r'='  
    RANGE = r'[A-Z][0-9]+:[A-Z][0-9]+'
    CELL_REFERENCE =r'([A-Z])([0-9]+)'
    NUMBER = r'\d+\.?(\d)*'
    FUNCTION_NAME =r'(SUM|AVERAGE|MIN|MAX)'
    CELL = r'[\' a-zA-Z_][ a-zA-Z0-9%)(_./-=;:+*-]*'
    SEMICOLON = r';'
    LPAREN = r'\('
    RPAREN = r'\)'
    PLUS = r'\+'
    MINUS = r'-'
    TIMES = r'\*'
    DIVIDE = r'/'
    PIPE = r'\|'
    
    # Ignored pattern
    ignore_newline = r'\n+'

    
    def ignore_newline(self, t):
        self.lineno += t.value.count('\n')

    def error(self, t):
        print(f"Illegal character '{t.value[0]}' at line {self.lineno}")
        self.index += 1

class SpreadsheetParser(Parser):
    tokens = SpreadsheetLexer.tokens

    precedence = (
        ('left', 'PLUS', 'MINUS'),
        ('left', 'TIMES', 'DIVIDE'),
        ('right', 'UMINUS')
    )

    def __init__(self):
        self.current_row = 1
        self.current_col = 'A'
        self.names = {}
        self.numeric_cells: Dict[str,float] = {}
        self.formula_cells: list[str] = []

    def _increment_column(self):
        # Convert current column letter to next letter (e.g., 'A' -> 'B', 'B' -> 'C')
        self.current_col = chr(ord(self.current_col) + 1)

    def _reset_column(self):
        self.current_col = 'A'
    
    def _increment_row(self):
        self.current_row += 1
        self._reset_column()

    def _get_cell_value(self, cell_id:str) -> float:
        return self.numeric_cells.get(cell_id, 0) 

    def _get_cell_id(self) -> str:
        return f"{self.current_col}{self.current_row}" 
    
    def _store_numeric_cell(self, cell_value:float):
        cell_id = self._get_cell_id()
        self.numeric_cells[cell_id]=cell_value
        print("I am in store_cell________Cell_id: ",cell_id)
        print("I am in store_cell:_______Cell value" ,cell_value)
        print("I am in store_cell________row number: ",self.current_row)
        self._increment_column()

    def _get_column_values(self, current_column:str, current_row:int, last_row:int) -> list[float]:
        """
        Get all numerical values in a column between two specified rows.
    
        Args:
            current_column (str): The starting column (e.g., 'A').
            current_row (int): The starting row (e.g., 1)
            last_row (int): The ending row (e.g., 5).
        
        Returns:
            list[float]: List of numerical values from the specified row range.
        """      
        column_values = []
        current_cell = f"{current_column}{current_row}"

        while current_row <= last_row:
            column_values.append(self._get_cell_value(current_cell))
            current_row += 1
            current_cell = f"{current_column}{current_row}"
    
        return column_values   

    def _get_row_values(self, current_column:str, current_row:int, last_column:str) -> list[float]:
        """
        Get all numerical values in a row between two specified columns.
    
        Args:
            current_column (str): The starting column (e.g., 'A').
            current_row (int): The starting row (e.g., 5)
            last_column (str): The ending column (e.g., 'E')
        
        Returns:
            list[float]: List of numerical values from the specified column range.
        """ 
        row_values = []
        current_cell = f"{current_column}{current_row}"

        while current_column <= last_column:  
            row_values.append(self._get_cell_value(current_cell))
            current_column = chr(ord(current_column) + 1)
            current_cell = f"{current_column}{current_row}"

        return row_values
    
    def _get_multiple_rows_columns_values(self, current_column:str, current_row:int, last_column:str, last_row:int) -> list[float]:
        """
        Get all numerical values in multiple rows and multiple columns between two specified columns and rows.
    
        Args:
            current_column (str): The starting column (e.g., 'A')
            current_row (int): The starting row (e.g., 1)
            last_column (str): The ending column (e.g., 'E')
            last_row (int): The ending row (e.g., 5)
        
        Returns:
            list[float]: List of numerical values from the specified rows and columns range.
        """ 
        result = []
        
        while current_column <= last_column:
            column_values = self._get_column_values(current_column, current_row, last_row)
            result.extend(column_values)  
            current_column = chr(ord(current_column) + 1)

        return result
    
    def _get_range_values(self, first_cell:str, last_cell:str):
        """
        Get numerical values from a specified range of cells in the spreadsheet.

        Args:
            first_cell (str): The starting cell reference of the range (e.g., 'A1').
            last_cell (str): The ending cell reference of the range (e.g., 'E5').
                    
        Returns:
            list[float]: A list of numerical values from the specified range.
        """ 
        first_column, first_row = self._parse_cell_reference(first_cell)
        last_column, last_row = self._parse_cell_reference(last_cell)

        if first_column == last_column:
            return self._get_column_values(first_column, first_row, last_row)
        
        elif first_row == last_row:
            return self._get_row_values(first_column,first_row, last_column)
        
        return self._get_multiple_rows_columns_values(first_column, first_row, last_column, last_row)
    
    def _parse_cell_reference(self, cell_reference: str):
        """
        Parse a cell reference and extract the column and row values.
    
        Args:
            cell_reference (str): The cell reference to parse  (e.g., 'A1', 'B12').
    
        Returns:
            Extracted column and row values.
        """
        match = re.match(SpreadsheetLexer.CELL_REFERENCE, cell_reference)
        column, row = match.groups()
        
        return column, int(row)
    
    @staticmethod
    def _calculate_function(func_name:str, operands: list[float]) -> float:
        """
        Performs spreadsheet function calculations on a list of numeric values.

        Args:
            func_name (str): The name of the function to execute (e.g 'SUM', 'AVERAGE', 'MIN', 'MAX').  
            operands (list[float]): List of numeric values to perform the calculation on.

        Returns:
            float: The result of the function calculation.
        """
        if func_name == 'SUM':
            return sum(operands)
        
        elif func_name == 'AVERAGE':
            return sum(operands) / len(operands)
        
        elif func_name == 'MIN':
            print("Operands_MIN: ", operands)
            return min(operands)
        
        elif func_name == 'MAX':
            print("Operands_MAX: ", operands)
            return max(operands)
        
        return 0  

    @_('header')
    def table(self, p):
        return (p.header, [])
    
    @_('header body')
    def table(self, p):
        return (p.header, p.body)
    
    @_('PIPE header_cells PIPE SEPARATOR')  
    def header(self, p):
        self._increment_row()
        self._increment_row()
        return p.header_cells
    
    @_('header_cells PIPE CELL')
    def header_cells(self, p):
        return p.header_cells + [p.CELL]
    
    @_('CELL')
    def header_cells(self, p):
        return [p.CELL]
    
    @_('body_row')
    def body(self, p):
        self._increment_row()
        return [p.body_row]
    
    @_('body body_row')
    def body(self, p):
        self._increment_row()
        return p.body + [p.body_row]
    
    @_('PIPE body_cells PIPE')
    def body_row(self, p):
        return p.body_cells
    
    @_('body_cells PIPE cell')
    def body_cells(self, p):
        return p.body_cells + [p.cell]
    
    @_('cell')
    def body_cells(self, p):
        return [p.cell]
    
    @_('CELL')
    def cell(self, p):
        self._increment_column()
        if p.CELL.startswith("'"):
            return p.CELL[1:]
        return p.CELL
    
    @_('EMPTY_CELL')
    def cell(self, p):
        self._increment_column() 
        return ""
       
    @_('formula')
    def cell(self, p):
        self.formula_cells.append(self._get_cell_id())
        self._store_numeric_cell(p.formula)   
        return p.formula
    
    @_('expr')
    def cell(self, p):
        self._store_numeric_cell(p.expr)
        return p.expr

    @_('EQUAL expr')
    def formula(self, p):
        return p.expr
    
    @_('EQUAL function')
    def formula(self, p):
        return p.function
    
    @_('EQUAL')
    def formula(self, p):
        return 0
    

    @_('FUNCTION_NAME LPAREN function_operands RPAREN')
    def function(self, p):
        if isinstance(p.function_operands, list):
            operands = p.function_operands
        else:
            operands = [p.function_operands]

        return self._calculate_function(p.FUNCTION_NAME, operands)
    
    @_('function_operands SEMICOLON expr')
    def function_operands(self, p):
        if isinstance(p.function_operands, list):
            return [p.expr] + p.function_operands 
        
        return [p.function_operands, p.expr]
        
    @_('function_operands SEMICOLON RANGE')
    def function_operands(self, p):
        first_cell, last_cell = p.RANGE.split(':')
        range_values = self._get_range_values(first_cell, last_cell)
        # Check if function_operands is a list or not
        if isinstance(p.function_operands, list):
             # If it's a list, extend it with the new range values
            return range_values + p.function_operands
              # If not a list, create a new list with the range values and the operand     
        return range_values + [p.function_operands] 
  
    @_('RANGE')
    def function_operands(self, p):
        first_cell, last_cell = p.RANGE.split(':')
        return self._get_range_values(first_cell, last_cell)
    
    @_('expr')
    def function_operands(self, p):
        return p.expr

    @_('expr PLUS expr')
    def expr(self, p):
        result = p.expr0 + p.expr1 
        return result

    @_('expr MINUS expr')
    def expr(self, p):
        result = p.expr0 - p.expr1
        return result

    @_('expr TIMES expr')
    def expr(self, p):
        result = p.expr0 * p.expr1
        return result

    @_('expr DIVIDE expr')
    def expr(self, p):
        result = p.expr0 / p.expr1
        return result

    @_('MINUS expr %prec UMINUS')
    def expr(self, p):
        return -p.expr

    @_('LPAREN expr RPAREN')
    def expr(self, p):
        return p.expr
        
    @_('CELL_REFERENCE')
    def expr(self, p):
        return self._get_cell_value(p.CELL_REFERENCE)
    
    @_('NUMBER')
    def expr(self, p):
        return float(p.NUMBER)
         
    def print_parsed_data(self,result):
        headers, body = result
    
        print("\nHeaders:")
        for i, header in enumerate(headers):
            print(f"  Column {chr(65 + i)}: {header}")

        print("\nRows:")
        for i, row in enumerate(body, start=3):  # Start numbering rows at 3
            print(f"  Row {i}:")
            for j, cell in enumerate(row):
                print(f"    Column {chr(65 + j)}: {cell}")

def main():
    try:
        filename = input("Enter the filename to read: ")
        
        with open(filename, 'r') as file:
            content = file.read()

        lexer = SpreadsheetLexer()
        parser = SpreadsheetParser()
        
        print("\nTokens in the file:")
        for token in lexer.tokenize(content):
            print(f"Type: {token.type}, Value: {token.value}, Line: {token.lineno}")

        result = parser.parse(lexer.tokenize(content))
        print("\nParsing result:", result)
        parser.print_parsed_data(result)

        html_content = HTMLGenerator.generate_table(result, parser.formula_cells)
        output_filename = filename.split('.')[0] +".html"
        with open(output_filename, 'w') as f:
            f.write(html_content)

        print(f"\nHTML table has been generated and saved to {output_filename}")
        print(parser.formula_cells)

    except FileNotFoundError:
        print(f"Error: The file '{filename}' does not exist.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()