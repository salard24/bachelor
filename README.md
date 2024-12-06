# Markdown Spreadsheet
Markdown Spreadsheet is an extension of the Markdown language that allows you to create spreadsheet-like tables and render them as HTML tables. It supports cell references and formula calculations, including functions like MIN, MAX, AVERAGE, and SUM, similar to the functionality found in spreadsheets like Excel and Google Sheets.\
Cell and separator are two important bulding blocks of each table.
1. Each cell should begin and end with a `|`.
2. A separator should begin and end with a `|` and inbetween should consist of minimum three `-`.

## Requirements
Markdown Spreadsheet requires the use of Python 3.6 or greater.

## Installation
1. Clone the repository  
2. Install required dependencies. Navigate to your project directory where `requirements.txt` is located and run:  

```bash
pip install -r requirements.txt
```
or
```bash
pip3 install -r requirements.txt
```
if you are using python3.

## The Table Structure
Each table consists of a single header row, which contains of header cells. A header cell can contain a string(empty string as well) with optinal numeric values. It should be noticed, that the header row should be followed by a separator. Here is an example:

```md
| Header1 |        |
|------------------|
```

Now, the table can be extended with body rows, which consist of body cells. A body cell can contain a string(empty string as well), a numeric value, or a formula. Furthermore, it is possible to create an empty cell by using the `EMPTY` string.\
A body row follows the same syntax as a header row, but it should not be followed by a separator. It should be noted that the body begins at row 3, it means the first cell reference would be `A3`. Here is an example:

```md
|  Header1  |  Header2  |
|-----------------------|
|   Cell    |   12345   |
|   EMPTY   |  Formula  |
```

## The Formula Structure
Each formula begins with an `=`, just like Excel and Google Sheets. There are five different types of formulas:
1. A cell reference can follow the `=`, for example, `= A3`
2. A numeric value can follow the `=`, for example, `= 55`
3. A function name can follow the `=`, for example, `= SUM(A3;55)`. In this case, the function requires some parameters which should be separated by a `;`. The parameters can be a cell reference, numeric value or range.
4. Parentheses can follow the `=`, for example, `=(55 + A3)`. This will allow to calculate expressions.
5. Finally, the use of a singel `=` is also possible, which means an empty formula, and its value is `0`.

If you want to have a formula as a string, withut calculating its value, you should put an `'`, before the formula. 

```md
|   Header1   |   Header2   |
|---------------------------|
|     10      |      20     |
| =A3 + 55    |  =55 + A3   |
| =SUM(A3;55) |  =(55 + A3) |
| =           |  =SUM(A3:A4)|
|   String    | '=SUM(A3:A4)|
```

## The Range Structure
Again just like other spreadsheets, it is possible to have a range as function's parameter. There are three kinds of ranges:
1. A range could be in the same column, for example, `A3:A5`
2. A range could be in the same row, for example, `A3:C3`
3. A range could be in diferent columns and rows, for example, `A3:C5`

## How to Use
The Markdown Spreadsheet can be executed by running the command:

```bash
python markdown_spreadsheet.py
```
or
```bash
python3 markdown_spreadsheet.py
```
if you are using python3.

Now the program will ask to give the desire .md file, in order to create the desire output, which is a html table with caluculated formula values.\
Below is a complete example of markdown file and its output: 

```md
|      Header1       |     Header2      |     Header3      |     Header4      |                     
|-----------------------------------------------------------------------------|
|         10         |        20        |        30        |      EMPTY       |
|         40         |        50        |        60        |      EMPTY       |
|        Sum         |    =SUM(A3:C3)   |    =SUM(A3:A4 )  |    =SUM(A3:C4)   |
|      Average       | =AVERAGE(A3:C3)  | =AVERAGE(A3:A4)  | =AVERAGE(A3:C4)  |
|        Min         |    =MIN(A3:C3)   |    =MAX(A3:A4)   |    =MAX(A3:C4)   |
|        Max         |    =SUM(A3:C3)   |    =SUM(A3:A4)   |    =SUM(A3:C4)   |
|Multiple parameters |  =SUM(A3:C4;A4)  |  =SUM(A3:C4;100) | =SUM(A3:C3;A4:C4)|  
```
Result:

```md
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
```
