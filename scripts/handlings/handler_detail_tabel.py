import sys
import os
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


from scripts.excel_reader import ExcelReader

class HandlerDetailTabel:
    def __init__(self):
        self.data = {}
    def main(self, path_to_detail_tabel:str, sheet_name: str = None, data: dict = {}):
        """
        return: Обычн
        """
        
        self.read_excel =  ExcelReader(path_to_detail_tabel, sheet_name=sheet_name)
        self.filter_data(path_to_detail_tabel)
        self.data.update(data)
        return self.data

    def filter_data(self, path):
        self.data =  self.read_excel.get_dict_all_data()
        dse = {}
        for row_num,row in self.data.items():
             
             if pd.isna(row.get('Unnamed: 0','')) or row_num < 7: continue
             dse[f'{os.path.basename(path)[:os.path.basename(path).index('.xls')]}_+_{row.get('Unnamed: 0','')}_++_{row.get('Unnamed: 3','')}_+++_{row_num}'] = {
                  'dse':row.get('Unnamed: 0','')
                  ,'detai' : row.get('Unnamed: 3','')
                  ,'count' : str(row.get('Unnamed: 8',''))
                }
        self.data = dse.copy()
        

if __name__ == '__main__':
    app = HandlerDetailTabel()
    data = app.main(path_to_detail_tabel=r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\ер6.212.112.xls")

    for row_num, row in data.items():
            print(row_num, row)