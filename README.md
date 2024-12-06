# Markdown Spreadsheet
Markdown Spreadsheet is an extension of the Markdown language that allows you to create spreadsheet-like tables and render them as HTML tables. It supports cell references and formula calculations, including functions like MIN, MAX, AVERAGE, and SUM, similar to the functionality found in spreadsheets like Excel and Google Sheets.\
Tables in this tool are built using two essential components:
1. **Cells:** Each cell should begin and end with a `|`.
2. **Separators:** This should begin and end with a `|` and contain at least three `-` characters in between.

## Requirements
* Markdown Spreadsheet requires the use of Python 3.6 or greater.

## Installation
1. Clone the repository   
2. Install the required dependencies. Navigate to your project directory where `requirements.txt` is located and run:  

```bash
pip install -r requirements.txt
```
or, if you are using python3:
```bash
pip3 install -r requirements.txt
```

## Table Structure
### Header Row
Each table consists of a single **header row**, which consists of header cells. A header cell can contain a string(including an empty string) with optional numeric values. The header row must always be followed by a separator. Here is an example:

```md
| Header1 |        |
|------------------|
```
### Body Rows
After the header, the table can include **body rows**, which consist of body cells. A body cell can contain:
* A string(including an empty string)
* A numeric value
* A formula

To create an empty cell, use the string `EMPTY`. Note that the **body begins at row 3**, so the first cell reference would be `A3`.\
Unlike the header row, body rows should not be followed by a separator. Here is an example:

```md
|  Header1  |  Header2  |
|-----------------------|
|   Cell    |   12345   |
|   EMPTY   |  Formula  |
```

## Formula Structure
Formulas in this tool work similarly to those in Excel and Google Sheets. Each formula must start with an `=`. The following types of formulas are supported:
1. **Cell References:** Example: `=A3`
2. **Numeric Values:** Example: `=55`
3. **Functions:** Example: `=SUM(A3;55)`
* Functions require parameters, separated by a  ` ;`. Parameters can be cell references, numeric values, or ranges.
4. **Expressions:** Example: `=(55 + A3)`
* Allows for calculations using parentheses.
5. **Empty Formula:** Example: `=`
* Results in a value of `0`

To include a formula as a string (without calculating its value), prefix it with `'`.\
Example: 

```md
|   Header1   |   Header2   |
|---------------------------|
|     10      |      20     |
| =A3 + 55    |  =55 + A3   |
| =SUM(A3;55) |  =(55 + A3) |
| =           |  =SUM(A3:A4)|
|   String    | '=SUM(A3:A4)|
```

## Range Structure
Ranges can be used as function parameters, similar to other spreadsheet tools:
1. **Single Column:** Example: `A3:A5`
2. **Single Row:** Example: `A3:C3`
3. **Multiple Columns and Rows:** Example: `A3:C5`

## How to Use
The Markdown Spreadsheet can be executed by running the command:

```bash
python markdown_spreadsheet.py
```
or, if you are using python3:
```bash
python3 markdown_spreadsheet.py
```

The program will prompt you to provide a `.md` file. It will then generate an HTML table with calculated formula values as the output.\
Below is a complete example of markdown file and its output: 

## Example Input Markdown File:
```md
|      Header1       |     Header2      |     Header3      |     Header4      |
|-----------------------------------------------------------------------------|
|         10         |        20        |        30        |      EMPTY       |
|         40         |        50        |        60        |      EMPTY       |
|        Sum         |    =SUM(A3:C3)   |    =SUM(A3:A4 )  |    =SUM(A3:C4)   |
|      Average       | =AVERAGE(A3:C3)  | =AVERAGE(A3:A4)  | =AVERAGE(A3:C4)  |
|        Min         |    =MIN(A3:C3)   |    =MIN(A3:A4)   |    =MIN(A3:C4)   |
|        Max         |    =MAX(A3:C3)   |    =MAX(A3:A4)   |    =MAX(A3:C4)   |
|Multiple parameters |  =SUM(A3:C4;A4)  |  =SUM(A3:C4;100) | =SUM(A3:C3;A4:C4)|
| Formula as string  |'=SUM(A3:C3;A4:C4)|  
```
## Example Output Table:

| Header1 | Header2 | Header3 | Header4 |
| --- | --- | --- | --- |
| 10.0 | 20.0 | 30.0 |     |
| 40.0 | 50.0 | 60.0 |     |
| Sum | 60.0 | 50.0 | 210.0 |
| Average | 20.0 | 25.0 | 35.0 |
| Min | 10.0 | 10.0 | 10.0 |
| Max | 30.0 | 40.0 | 60.0 |
| Multiple parameters | 250.0 | 310.0 | 210.0 |
| Formula as string | \=SUM(A3:C3;A4:C4) | 

