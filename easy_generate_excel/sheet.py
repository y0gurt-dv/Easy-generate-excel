from typing import Union, Literal, List
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter, get_column_interval
from copy import copy


class Sheet:
    def __init__(
            self,
            name: str,
            headers: list,
            data: List[list],
            bold_header: bool = True,
            auto_filter: bool = True,
            center_cols_indexes: Union[Literal['all'], List[int]] = 'all',
            center_headers_indexes: Union[Literal['all'], List[int]] = 'all'):
        self.name = name
        self.headers = headers
        self.data = data
        self.auto_filter = auto_filter
        self.bold_header = bold_header
        self.center_cols_indexes = center_cols_indexes
        self.center_headers_indexes = center_headers_indexes

    @property
    def max_row(self):
        return len(self.data) + 1

    @property
    def max_col(self):
        return len(self.headers)

    @property
    def cols_interval(self):
        return get_column_interval(1, self.max_col)

    @property
    def full_data(self):
        data = list(self.data)
        data.insert(0, self.headers)
        return data

    def __repr__(self) -> str:
        return f'<Sheet name:{self.name}>'

    def _need_center_col(self, col_idx: int, row_idx: int) -> bool:
        if row_idx == 0:
            indexes_list = self.center_headers_indexes
        else:
            indexes_list = self.center_cols_indexes

        if indexes_list == 'all':
            return True
        return col_idx in indexes_list

    def to_dict(self):
        return {
            'name': self.name,
            'headers': self.headers,
            'data': self.data,
            'auto_filter': self.auto_filter,
            'bold_header': self.bold_header,
            'center_cols_indexes': self.center_cols_indexes,
            'center_headers_indexes': self.center_headers_indexes,
        }

    @classmethod
    def from_dict(cls, raw: dict):
        new_instance = cls(
            name=raw['name'],
            headers=raw['headers'],
            data=raw['data'],
            auto_filter=raw.get('auto_filter', True),
            bold_header=raw.get('auto_filter', True),
            center_cols_indexes=raw.get('center_cols_indexes', 'all'),
            center_headers_indexes=raw.get('center_headers_indexes', 'all'),
        )
        return new_instance

    def create(self, workbook: Workbook):
        worksheet = workbook.create_sheet(self.name)
        max_rows_widths = {idx: 0 for idx in self.cols_interval}
        data = self.full_data
        cells: List[List[Cell]] = worksheet.iter_rows(
            min_row=1,
            min_col=1,
            max_row=self.max_row,
            max_col=self.max_col
        )

        for row in cells:
            for cell in row:
                row_idx = cell.row - 1
                col_idx = cell.column - 1
                col_letter = cell.column_letter

                value = str(data[row_idx][col_idx])
                width = len(max(value.split('\n'), key=len))

                cell.value = value
                max_rows_widths[col_letter] = max(
                    max_rows_widths[col_letter],
                    width
                )

                if row_idx == 0 and self.bold_header:
                    new_font = copy(cell.font)
                    new_font.bold = True
                    cell.font = new_font

                if self._need_center_col(col_idx, row_idx):
                    cell.alignment = Alignment(
                        vertical='center',
                        horizontal='center'
                    )

        for letter, width in max_rows_widths.items():
            if width != 0:
                worksheet.column_dimensions[letter].width = max(
                    (width + 2) * 1.2,
                    9
                )

        if self.auto_filter:
            worksheet.auto_filter.ref = 'A1:{}1'.format(
                get_column_letter(self.max_col)
            )
