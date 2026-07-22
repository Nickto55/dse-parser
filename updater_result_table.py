import pandas as pd
from scripts.excel_reader import ExcelReader
# from scripts.excel_enter import ExcelDataInserter

class UpdaterResultTableLogic:
    def __init__(self,file_path_main, file_path_cz):
        self.file_path_main = file_path_main
        self.file_path_cz = file_path_cz[0]

        self.data_main_processing = {}
        self.data_cz_sorted = {}
        self.count_del = 0
        self.data_del = {}
        self.main_data={}
        self.healders = []

    def main(self):
        self.count_del = 0
        self.data_del = {}
        print(f'file_path_main: {self.file_path_main}')
        print(f'file_path_cz: {self.file_path_cz}')
        reader = ExcelReader(self.file_path_main, sheet_name='Изделия')
        self.main_data = reader.get_dict_all_data()

        reader = ExcelReader(self.file_path_cz, sheet_name='Лист1')

        self.filter_cz_data(reader.get_dict_all_data())
        self.filter_main_data()



        for num_row, data_row in self.data_main_processing.items():
            dse_reply = data_row.get('ДСЕ (изделия)', '').upper()
            count_dse = 0
            for num_row_search, data_row_search in self.data_main_processing.items():
                if dse_reply == data_row_search.get('ДСЕ (изделия)', '').upper():
                    count_dse += 1
            self.data_main_processing[num_row]['частота упоминаний'] = count_dse

        print(f'Было: {len(self.main_data):<6}', f'Сумма ост.+вык.: {sum([len(self.data_main_processing), self.count_del])}',
              f'Остал: {len(self.data_main_processing)}', f'del: {self.count_del}', f'Уник дсе del: {len(self.data_del)}')

        return self.data_main_processing, self.healders

    def filter_cz_data(self, cz_data):
        data_cz_processing = {}
        for num_row, row_dse in cz_data.items():
            dse = row_dse.get('ДСЕ', '').upper()
            if not dse in data_cz_processing.keys():
                data_cz_processing.update({f'{dse}': False})

            if pd.isna(row_dse.get('Выполнено')) and pd.isna(row_dse.get('Закрыто')):
                data_cz_processing[dse] = True

        self.data_cz_sorted =  data_cz_processing

    def filter_main_data(self):
        keys_cz = list(self.data_cz_sorted.keys())

        # print(keys_cz)
        for i, row_dse in self.main_data.items():
            dse = row_dse.get('ДСЕ (детали)', '').upper()
            if dse in keys_cz:
                if self.data_cz_sorted.get(dse, ''):
                    self.data_main_processing.update({len(self.data_main_processing): row_dse})
                else:
                    self.data_del.update({dse: f'{self.data_cz_sorted.get(dse, "")}'})
                    self.count_del += 1
            else:
                print(f'Ошибка дсе: {dse} нет в СЗ')

        self.healders = list(row_dse.keys())





if __name__=='__main__':
    app = UpdaterResultTableLogic(
        file_path_main=r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\2_logic_prog\__Переводы по изделиям, годовой план 2026,07,15 — копия.xlsx"
        , file_path_cz=r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\2_logic_prog\ДСЕ по СЗ и Извещениям.xlsx"
    )

    app.main()

# excel_printer = ExcelDataInserter(file_path_main)
# excel_printer.insert_data({'': self.data_main_processing}, sheet_name='Изделия', headers=list(row_dse.keys()))


