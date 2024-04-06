import pandas as pd
import os
import shutil
from openpyxl.styles import PatternFill, Border, Side
from datetime import datetime

class ExcelHandler:
    def __init__(self, excel_filepath: str) -> None:
        if not os.path.exists(excel_filepath):
            raise FileNotFoundError(f'Nie mozna znalexc pliku: {excel_filepath}')
        self.excel_filepath = excel_filepath
        self.csv_dir = 'csv_data/'
        self.polish_months_names = {
                    1: "Styczeń",
                    2: "Luty",
                    3: "Marzec",
                    4: "Kwiecień",
                    5: "Maj",
                    6: "Czerwiec",
                    7: "Lipiec",
                    8: "Sierpień",
                    9: "Wrzesień",
                    10: "Październik",
                    11: "Listopad",
                    12: "Grudzień",
                }

        self.CELL_COLOR = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        self.BORDER = Side(style='thin', color='000000')
        self.BORDER_STYLE = Border(left=self.BORDER, right=self.BORDER, top=self.BORDER, bottom=self.BORDER)
        
    def get_data_to_csv_files(self, column_pjo: str = 'pjo', column_desc: str= 'opis', column_income: str = 'przychód', column_expenditure: str = 'rozchód') -> None:
        if os.path.exists(self.csv_dir):
            shutil.rmtree(self.csv_dir)
        os.mkdir(self.csv_dir)
        
        DATA = pd.read_excel(self.excel_filepath)
        PJO = list(filter(lambda x: isinstance(x, str), set(DATA[column_pjo])))
        PJO.remove('nie kopiować')
        PJO_DICT = {j: {} for j in PJO}

        for _, row in DATA.iterrows():
            income = row[column_income]
            expenditure = row[column_expenditure]
            
            if row[column_pjo] in PJO:
                PJO_DICT[row[column_pjo]][row[column_desc]] = (income, expenditure)
        
        for p in PJO:
            with open(f'{self.csv_dir}/{p}.csv', 'a') as file:
                file.write('description;income;expenditure\n')
                for desc, money in PJO_DICT[p].items():
                    income, expend = money
                    file.write(f'{desc};{income:.2f};{expend:.2f}\n')
    
    def create_excel_file(self, output_filename: str = 'output.xlsx') -> None:
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            for file in os.listdir(self.csv_dir):
                filename = file.removesuffix('.csv')
                data = pd.read_csv(f'{self.csv_dir}{file}', sep=';')
                data.to_excel(writer, sheet_name=filename, index=False)

                sheet = writer.sheets[filename]
                sheet.cell(1, 4, 'SUMA') # description of value below
                sheet.cell(1, 4).fill = self.CELL_COLOR
                sheet.cell(1, 4).border = self.BORDER_STYLE
                end_cell = len(data['description']) + 2
                sheet.cell(2, 4, f'=SUM(B2:B{end_cell})-SUM(C2:C{end_cell})')
                sheet.cell(2, 4).fill = self.CELL_COLOR

    def write_data_to_one_excel_file(self, output_filename: str = 'output.xlsx') -> None:
        if not os.path.exists(output_filename):
            self.create_excel_file(output_filename)
            return
        with pd.ExcelFile(output_filename, engine='openpyxl') as writer:
            for file in os.listdir(self.csv_dir):
                filename = file.removesuffix('.csv')
                excel_data = pd.read_excel(writer, sheet_name=filename)
                data_length = len(excel_data['description'])
                print(filename, data_length)
                sheet = writer.sheets[filename]
                month = self.polish_months_names[datetime.now().month]
                sheet.cell(1, data_length, month)
                sheet.cell(1, data_length).fill = self.CELL_COLOR

                



if __name__ == '__main__':
    filepath = input('filename: ')
    test = ExcelHandler(filepath)
    #test.get_data_to_csv_files()
    test.write_data_to_one_excel_file()
