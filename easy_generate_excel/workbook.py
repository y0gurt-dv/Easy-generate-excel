import os
from io import BytesIO
from typing import List, Optional, Union

from openpyxl import Workbook

from .sheet import Sheet


class ExcelFile:
    def __init__(self, sheets: Union[List[dict], List[Sheet]]):
        self.sheets = sheets

    @property
    def sheets(self) -> List[Sheet]:
        """
            Return list of sheets
        """
        return list(self._sheets.values())

    @sheets.setter
    def sheets(self, sheets: Union[List[dict], List[Sheet]]):
        self.clear_sheets()
        for sheet in sheets:
            self.add_sheet(sheet)

    def add_sheet(self, sheet: Union[dict, Sheet]):
        """
            Add sheet
        """
        if isinstance(sheet, dict):
            sheet = Sheet.from_dict(sheet)

        if self.get_sheet(sheet.name):
            raise Exception(
                'Sheet with name {} already exist'.format(
                    sheet.name
                )
            )

        self._sheets.setdefault(
            sheet.name,
            sheet
        )

    def get_sheet(self, name: str) -> Optional[Sheet]:
        """
            Return sheet by name
        """
        return self._sheets.get(name, None)

    def delete_sheet(self, name: str) -> Optional[Sheet]:
        """
            Delete sheet by name
        """
        return self._sheets.pop(name, None)

    def clear_sheets(self):
        """
            Delete all sheets
        """
        self._sheets = {}

    def create(self, return_bytes: bool = False) -> Union[BytesIO, Workbook]:
        """
            Create new Workbook with all sheets\n
            Return BytesIO or WorkBook
        """
        workbook = Workbook()
        del workbook['Sheet']

        for sheet in self.sheets:
            sheet.create(workbook=workbook)

        if return_bytes:
            io_bytes = BytesIO()
            workbook.save(io_bytes)

            return io_bytes

        return workbook

    def create_file(self, output_name: str, output_path: str) -> str:
        """
            Create and save new file with ``output_name`` to ``output_path``\n
            ``output_name`` must be ``WITHOUT`` format
        """
        bytes: BytesIO = self.create()
        filepath = os.path.join(output_path, f'{output_name}.xlsx')
        with open(filepath, 'wb') as fp:
            fp.write(bytes.getvalue())

        return filepath

    def to_dict(self):
        """
            Return config dict
        """
        return [
            sheet.__dict__
            for sheet in self.sheets
        ]
