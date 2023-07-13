# python-openpyxl

* **Workbook:** A spreadsheet is represented as a workbook in openpyxl. A workbook consists of one or more sheets.
* **Sheet:** A sheet is a single page composed of cells for organizing data.
* **Cell:** The intersection of a row and a column is called a cell. Usually represented by A1, B5, etc.
* **Row:** A row is a horizontal line represented by a number (1,2, etc.).
* **Column:** A column is a vertical line represented by a capital letter (A, B, etc.).

```
sheet["A1"].value
sheet.cell(row=10, column=6).value

sheet["A"]
sheet["A1:C2"]
sheet["A:B"]
```

There are also multiple ways of using normal Python [generators](https://realpython.com/introduction-to-python-generators/) to go through the data. The main methods you can use to achieve this are:
```
* .iter_rows()
* .iter_cols()
```

Both methods can receive the following arguments:
```
* min_row
* max_row
* min_col
* max_col
```

Links
* https://www.javatpoint.com/python-openpyxl
* https://realpython.com/openpyxl-excel-spreadsheets-python/
