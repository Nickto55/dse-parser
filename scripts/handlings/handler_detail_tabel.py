import sys
import os
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from scripts.excel_reader import ExcelReader


class HandlerDetailTabel:
    def __init__(self):
        self.data = {}
        self.dse_product = ''

    def main(self, path_to_detail_tabel: str, sheet_name: str = None, data: dict = {}):
        """
        return: Обычн
        """

        self.read_excel = ExcelReader(path_to_detail_tabel, sheet_name=sheet_name)
        self.filter_data(path_to_detail_tabel)
        self.data.update(data)
        self.count_dse_detail()
        return self.data

    def count_dse_detail(self):
        data_count = {}
        dict_data_key = {}

        for keys_for_data in self.data.keys():
            dse = keys_for_data[keys_for_data.index('_+_') + 3:keys_for_data.index('_++_')]
            if dse in dict_data_key.keys():
                dict_data_key[dse].append(keys_for_data)
            else:
                dict_data_key[dse]=[keys_for_data]

        for dse_product, row_detail in self.data.items():
            dse = row_detail.get('dse', '')
            if dse in data_count.keys():
                data_count[dse] = int(data_count.get(dse, '')) + 1
            else:
                    data_count[dse] = 1


        for key_detal, value_detal in data_count.items():
            for key_for_dse_data in dict_data_key.get(key_detal):
                self.data[key_for_dse_data].update({"count dse": value_detal})

    def filter_data(self, path):
        self.data = self.read_excel.get_dict_all_data()
        dse = {}
        self.dse_product = os.path.basename(path)[:os.path.basename(path).index('.xls')]
        for row_num, row in self.data.items():
            if pd.isna(row.get('Unnamed: 0', '')) or row_num < 7: continue

            dse[f'{self.dse_product}_+_{row.get('Unnamed: 0', '')}_++_{row.get('Unnamed: 3', '')}_+++_{row_num}'] = {
                'dse': row.get('Unnamed: 0', '')
                , 'detai': row.get('Unnamed: 3', '')
                , 'count': str(row.get('Unnamed: 8', ''))
            }
        self.data = dse.copy()


if __name__ == '__main__':
    app = HandlerDetailTabel()
    data = app.main(path_to_detail_tabel=r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\ер6.212.112.xls")

    for row_num, row in data.items():
            print(row_num, row)
