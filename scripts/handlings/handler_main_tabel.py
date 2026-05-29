import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


from scripts.excel_reader import ExcelReader

class HandlerProductTabel:
    def __init__(self):
        self.data = ()
    def main(self, path_to_main_tabel:str):
        """
        return: Обычн
        """
        self.read_excel =  ExcelReader(path_to_main_tabel)
        self.filter_data()
        return self.data

    def filter_data(self):
        self.data =  self.read_excel.get_dict_all_data()
        dse = {}
        for row_num,row in self.data.items():
             dse[str(row.get('ДСЕ','')).lower()] = {
                  'name' : row.get('Имя','')
                  ,'rc' : str(row.get('РЦ',''))
                }
        self.data = dse.copy()
        

if __name__ == '__main__':
    app = HandlerProductTabel()
    data = app.main(path_to_main_tabel=r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\Копия Список переводов 28.05.xlsx")

    for row_num, row in data.items():
            print(row_num, row)