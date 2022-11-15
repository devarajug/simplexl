# SIMPLEXL

Simplexl is a Python package which is used to create excel files dynamically using a program. This package depends on [openpyxl](https://pypi.org/project/openpyxl/) which is native python package for creating excel.

# FEATURES

- Generated formatted excel as ouput.
- Automatically pick the width of column based on data size of column.


# Installation
```
pip install simplexl
```

# How to use

The usage of simplexl is as follows 
- 

```
from simplexl import CreateExcel

xl = CreateExcel()


xl.create_sheet(
    sheet_name=sheet_name,   # optional
    sheet_index=sheet_index  # optional default = 0
    col_data=col_data,
    row_data=row_data
)

xl.save(
    name="path/name of the excel  #  default = generate-simplexl.xlsx
)
```

# Example

```
from simplexl import CreateExcel

sheet1_col_data = ['col1', 'col2']
sheet1_row_data = [
    ('col1_row1', 'col2_row1'),
    ('col1_row2', 'col2_row2')
]

sheet2_col_data = ['col1', 'col2']
sheet2_row_data = [
    ('col1_row1', 'col2_row1'),
    ('col1_row2', 'col2_row2')
]

xl = CreateExcel()

# Sheet 1  
xl.create_sheet(
    col_data=sheet1_col_data,
    row_data=sheet1_row_data
)

# Sheet 2
xl.create_sheet(
    col_data=sheet2_col_data,
    row_data=sheet2_row_data
)
xl.save()
```
It will create a xlsx file using row and col data 

# License

Copyright (c) 2022 Devaraju Garigapati

This repository is licensed under the [MIT](https://opensource.org/licenses/MIT) license.
See [LICENSE](https://opensource.org/licenses/MIT) for details.