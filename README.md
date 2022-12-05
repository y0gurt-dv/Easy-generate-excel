# Easy generate excel

An easy way to create an excel file with data

# Overview

```python
from easy_generate_excel import ExcelWorkBook, Sheet
```

### Sheet

Sheet config class

###### Parameters

| Name                       | Required | Default    | Type                            | Description                                               |
| -------------------------- | -------- | ---------- | ------------------------------- | --------------------------------------------------------- |
| name                       | True     | -          | str                             | Sheet name                                                |
| headers                    | True     | -          | list                            | Headers list                                              |
| data                       | True     |            | list[list]                      | Data list. Data will be read by data[row_idx][col_idx]    |
| bold_header                | False    | True       | bool                            | Need bold header                                          |
| auto_filter                | False    | True       | bool                            | Need enable auto filter in headers                        |
| center_cols_indexes        | False    | all        | Literal['all'], list[int], None | Cols indexes of data cols which will be center            |
| center_headers_indexes     | False    | all        | Literal['all'], list[int], None | Cols indexes of headers cols which will be center         |
| not_center_cols_indexes    | False    | None       | Literal['all'], list[int], None | Cols indexes of data cols which will ``NOT`` be center    |
| not_center_headers_indexes | False    | None       | Literal['all'], list[int], None | Cols indexes of headers cols which will ``NOT`` be center |
| cell_expansion             | False    | int, float | 1.2                             | The percentage by which the cell width will be increased  |
| min_cell_width             | False    | int        | 9                               | Min cell width (Excel default 9)                          |

##### 

### ExcelWorkBook

The main class for generating excel

###### Parameters

| Name   | Required | Default | Type                    | Description            |
| ------ | -------- | ------- | ----------------------- | ---------------------- |
| sheets | True     | -       | List[dict], List[Sheet] | List of sheets configs |

You can pass an array of dictionaries, they will be converted to ``Sheet``





# Installation

```bash
pip install easy_generate_excel
```

```python
from easy_generate_excel import ExcelWorkBook, Sheet
```



# Example

```python
from easy_generate_excel import ExcelWorkBook, Sheet
from io import BytesIO
from openpyxl import Workbook

sh = Sheet(
    name='Sheet 1',
    headers=[
        'Test header 1',
        'Test header 2',
        'Test header 3',
        'Test header 4',
        'Test header 5',
        'Test header 6',
    ],
    data=[
        ['Data 1', 'Data 2', 'Data 3', 'Data 4', 'Data 5', 'Data 6'],
        ['Data 7', 'Data 8', 'Data 9', 'Data 10', 'Data 11', 'Data 12'],
        ['Data 13', 'Data 14', 'Data 15', 'Data 16', 'Data 17', 'Data 18'],
        ['Data 19', 'Data 20', 'Data 21', 'Data 22', 'Data 23', 'Data 24'],
    ],
    # NOT REQUIRED
    bold_header=True,
    auto_filter=True,
    center_cols_indexes='all',
    center_headers_indexes='all',
    not_center_cols_indexes=None,
    not_center_headers_indexes=None,
    cell_expansion=1.2,
    min_cell_width=9,
)
sh2 = Sheet(
    name='Sheet 2',
    headers=[
        'Test header 1',
        'Test header 2',
        'Test header 3',
    ],
    data=[
        ['Data 1', 'Data 2', 'Data 3'],
        ['Data 4', 'Data 5', 'Data 6'],
        ['Data 7', 'Data 8', 'Data 9',],
        ['Data 10', 'Data 11', 'Data 12'],
        ['Data 13', 'Data 14', 'Data 15',],
        ['Data 16', 'Data 17', 'Data 18'],
        ['Data 19', 'Data 20', 'Data 21',],
        ['Data 22', 'Data 23', 'Data 24'],
    ],
    # NOT REQUIRED
    bold_header=True,
    auto_filter=True,
    center_cols_indexes='all',
    center_headers_indexes='all',
    not_center_cols_indexes=None,
    not_center_headers_indexes=None,
    cell_expansion=1.2,
    min_cell_width=9,
)

workbook_factory = ExcelWorkBook(sheets=[sh, sh2])

# Return BytesIO
output_bytes: BytesIO = workbook_factory.create(return_bytes=True)
with open('test_file.xlsx', 'wb') as fp:
    fp.write(output_bytes.getvalue())

# Return workbook
output_workbook: Workbook = workbook_factory.create()
output_workbook.save('test_file.xlsx')

# Save file
output_filepath: str = workbook_factory.create_file(
    output_name='test_file',
    output_path=''
)

```


