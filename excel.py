from numpy import intp
import pandas as pd
import os
import shutil
from openpyxl.styles import PatternFill, Border, Side

class ExcelHandler:
    def __init__(self, excel_filepath: str) -> None:
        if not os.path.exists(excel_filepath):
            raise FileNotFoundError(f'Nie mozna znalexc pliku: {excel_filepath}')
        self.excel_filepath = excel_filepath
        self.csv_dir = 'csv_data/'
        
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
    
filepath = input('asd: ')
test = ExcelHandler(filepath).get_data_to_csv_files()
