import pandas as pd


class ExcelReader:
    def __init__(self, xlsx_path):
        self.xlsx_data = pd.read_excel(xlsx_path)
        print("excel reader init")

    def __repr__(self):
        return f"xlsx data: {self.xlsx_data}"

    def rows(self):
        return self.xlsx_data.shape[0]

    def columns(self):
        return self.xlsx_data.shape[1]

    def get_content(self, row, title):
        if row < 0 or row > self.rows():
            return None

        for i in range(self.columns()):
            if self.xlsx_data.columns[i] == title:
                return self.xlsx_data.loc[row][i]

        return None
