class HTMLGenerator:
    """Generates HTML table representation of spreadsheet data with styling and cell highlighting."""  

    CSS_STYLES_SPREADSHEET="""
            table {
                border-collapse: collapse;
                font-family: 'Calibri', sans-serif;
                font-size: 11pt;
                width: 100%;
                background-color: white;
            }
            th {
                background-color: #f3f4f6;
                border: 1px solid #e5e7eb;
                padding: 12px;
                font-weight: 600;
                text-align: left;
                color: #374151;
            }
            th:hover {
                background-color: #e5e7eb;
            }
            td {
                border: 1px solid #e5e7eb;
                padding: 12px;
                color: #4b5563;
            }
            tr:nth-child(even) {
            background-color: #f9fafb;
            }
            tr:hover {
                background-color: #f3f4f6;
            }
            .highlighted-cell {
            background-color: #fef3c7;
            }"""

    @classmethod
    def _create_html_head(cls):
        """Creates HTML head section with embedded styles."""

        return f"""    <head>
        <style>{cls.CSS_STYLES_SPREADSHEET}
        </style>
    </head>"""

    @staticmethod
    def _create_table_header(headers):
        """Generates HTML for table header row."""

        header_cells = "\n".join([f"                <th>{header}</th>" for header in headers])
        return f"    <tr>\n{header_cells}\n            </tr>"

    @staticmethod
    def _create_table_body(body, highlighted_cells):
        """Generates HTML for table body with optional cell highlighting."""

        highlighted_coordinates = HTMLGenerator._get_highlighted_coordinates(highlighted_cells)
        rows = []
        
        for row_idx, row in enumerate(body):
            cells = []
            for col_idx, cell in enumerate(row):
                highlight_class = ""
                if (row_idx + 3, col_idx) in highlighted_coordinates:  # +3 because data starts from row 3
                    highlight_class = ' class="highlighted-cell"'
                cells.append(f'                <td{highlight_class}>{cell}</td>')    

            cells_html = "\n".join(cells)
            rows.append(f"    <tr>\n{cells_html}\n            </tr>")           

        return "        \n        ".join(rows)

    @classmethod
    def generate_table(cls, result, highlighted_cells=None):
        """
        Generates complete HTML table from spreadsheet data.
        
        Args:
            result: Tuple of (headers, body) containing spreadsheet data
            highlighted_cells: Optional list of cell IDs to highlight
        
        Returns:
            String containing complete HTML document
        """
        headers, body = result
        
        return f"""<!DOCTYPE html>
<html>
{cls._create_html_head()}
    <body>
        <table>
        {cls._create_table_header(headers)}
        {cls._create_table_body(body, highlighted_cells)}
        </table>
    </body>
</html>"""

    @staticmethod
    def _get_cell_coordinates(cell_id):
        """Converts spreadsheet cell ID (e.g., 'A1') to (row, column) coordinates."""

        col = ord(cell_id[0]) - ord('A')
        row = int(cell_id[1:]) 

        return row, col
    
    @staticmethod
    def _get_highlighted_coordinates(highlighted_cells):
        """Converts list of cell IDs to set of (row, column) coordinates for highlighting."""
        
        coordinates = set()
        for cell_id in highlighted_cells:
            row, col = HTMLGenerator._get_cell_coordinates(cell_id)
            coordinates.add((row, col))

        return coordinates