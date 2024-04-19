import gspread
import re
import string

from google.oauth2.service_account import Credentials
from gspread_formatting import *

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "1gP0-ZDDa2I8V4drA2XxhIAbHtpGB4BE_KM8nFNhWPvc"

workbook = client.open_by_key(sheet_id)

class Table:
    def __init__(self, title, cj, constraints):
        self.title = title
        self.constraints = constraints

        # To keep track on how many slacks were added
        self.added_zeros = 0

        # To keep track of the original count of Cjs
        self.cj_count = len(cj)
        self.cj = ["", "", "Cj"]

        # Add the given cjs
        for i in range(len(cj)):
            self.cj.append(str(cj[i]))

        # Create the additional slacks
        for _ in range(len(constraints)):
            self.cj.append("0")
            self.added_zeros += 1

        # The total row or constraints tin the table
        self.row_count = len(self.cj) - len(cj) - 3

        # For designing purposes
        self.cj_row = workbook.sheet1.find("Cj").row
        self.cj_zj_row = workbook.sheet1.find("Cj - Zj").row
        self.solution_row = workbook.sheet1.find("SOLUTION").row

    def create_column_names(self):
        # Create Column Names
        column_names = ["", "CV", "BV", "X", "Y"]
        for i in range(self.added_zeros):
            column_names.append(f"S{i + 1}")
        column_names.append("SOLUTION")
        column_names.append("RATIO")

        return column_names


    def create_rows(self):
        rows = []
        for i in range(self.row_count):
            row = [f"R{i + 1}", "0", f"S{i + 1}"]
            for c in constraints[i][:-1]:
                row.append(str(c))

            for j in range(self.row_count):
                row.append(f"{'1' if i == j else '0'}")

            row.append(self.constraints[i][len(self.constraints[i]) - 1])

            rows.append(row)

        return rows



    def get_zj(self):
        zj = ["", "", "Zj"]
        for _ in range(self.row_count + self.cj_count):
            zj.append("0")
        zj.append("Z0=0")
        return zj


    def cj_zj(self, zj):
        cj_zj_list = ["", "", "Cj - Zj"]
        cj_list = [x for x in self.cj if re.findall(r"\d", x)]
        zj_list = [x for x in zj if re.findall(r"\d", x)][:-1]

        for i in range(len(cj_list)):
            total = int(cj_list[i]) - int(zj_list[i])
            cj_zj_list.append(total)

        cj_zj_list.append("")
        cj_zj_list.append("END")

        return cj_zj_list


    def styling(self):



        wk = workbook.sheet1

        table_row_start = wk.find(self.title).row
        table_col_end, table_row_end = wk.find("RATIO").col, wk.find("RATIO").row
        cv_row = wk.find("CV").row

        # Convert column number into alphabet
        columns = string.ascii_uppercase
        table_end_column = columns[table_col_end - 1]

        main_bg = (0.9764705882352941, 0.796078431372549, 0.611764705882353)

        general_format = CellFormat(
            textFormat=TextFormat(fontSize=13),
            horizontalAlignment="CENTER"
        )

        fmt = CellFormat(
            backgroundColor=Color(main_bg[0], main_bg[1], main_bg[2]),
            textFormat=TextFormat(bold=True, foregroundColor=Color(0, 0, 0), fontSize=13),
            horizontalAlignment="CENTER"
        )


        format_cell_ranges(wk, [
            (f'B{cv_row}:J{cv_row}', fmt),
            (f'C{self.cj_row}:C{self.cj_zj_row}', fmt),
            (f"A{table_row_start}:{table_end_column}:{table_row_end}", general_format),
            (f"I{self.solution_row + self.added_zeros + 1}", fmt)
        ])

    def draw(self):

        workbook.sheet1.clear()

        # Generate Column Names to be printed
        column_names = self.create_column_names()

        # Generate Zjs
        zj = self.get_zj()

        cols_range = 10
        drawing = [
            [self.title, "", "", "", "", "KEY COLUMN", "KEY ROW", "KEY ELEMENT", "VALUE <= 0"],
            ["" for _ in range(cols_range)],
            [n for n in self.cj],
            column_names,
        ]

        # Draws each row
        rows = self.create_rows()
        for row in rows:
            drawing.append(row)

        # Draws the ZJs
        drawing.append(zj)

        # Draws the CJ-ZJ
        cj_zj = self.cj_zj(zj)
        drawing.append(cj_zj)

        workbook.sheet1.update("A1", drawing)
        self.styling()





cj = [3, 5]
constraints = [
    [6, 4, 35],
    [8, 2, 20],
    [1, 6, 20]
]


# cj = [100, 85]
# constraints = [
#     [3, 2, 48],
#     [1, 2, 18],
# ]

table = Table("Table 0", cj, constraints)

table.draw()