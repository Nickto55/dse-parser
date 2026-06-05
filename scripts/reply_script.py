import os.path

import pandas as pd

from scripts.handlings.handler_reply import HandlerReplyTabel

class ScriptReplyTabel:
    def __init__(self):
        self.data = {}
        self.receive_data = {}
        self.data_handler = {}
        self.path_file = None

    def main(self, path_reply_tabl_file, receive_data: dict):
        self.receive_data = receive_data.copy()
        self.path_file = path_reply_tabl_file

        handler_reply_file = HandlerReplyTabel()
        self.data_handler =  handler_reply_file.main(path_to_reply_tabel=path_reply_tabl_file).copy()

        self.filter_for_data()


        return self.data

    def filter_for_data(self):
        list_of_dse_and_name_from_handler_file = []
        self.data = self.receive_data.copy()

        for num_row, row_data_handler in self.data_handler.items():
            dse_and_name_halding_file = row_data_handler.get('Unnamed: 4','')
            if pd.isna(dse_and_name_halding_file): continue
            list_of_dse_and_name_from_handler_file.append(dse_and_name_halding_file)

        for global_key,global_value in self.receive_data.items():
            for row_key, row_data in global_value.items():
                dse_and_name= f"{row_data.get('ДСЕ (изделия)','')} {row_data.get('Наименование (изделия)','')}"
                dse = row_data.get('ДСЕ (изделия)','')
                row_return = row_data.copy()

                row_return[os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]] = f''
                row_return[f"Dse {os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]}"] = f''

                if dse_and_name in list_of_dse_and_name_from_handler_file:

                    row_return[os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]] = '+'

                else:
                    for i in list_of_dse_and_name_from_handler_file:
                        if dse_and_name in i:
                            row_return[os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]] = f'+-'
                            row_return[f"Dse {os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]}"]=f'{i}'
                            self.data[global_key][row_key] = row_return.copy()
                        elif dse in i:
                            if '-' in i[:i.index(' ')] and not '-' in dse: continue
                            if '-' in dse:
                                if '.' in  i[i.index('-'):i.index(' ')] and not '.' in dse[dse.index('-'):]: continue
                            row_return[os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]] = f'+--'
                            row_return[f"Dse {os.path.basename(self.path_file)[:os.path.basename(self.path_file).index('.xl')]}"]=f'{i}'
                            self.data[global_key][row_key] = row_return.copy()
                self.data[global_key][row_key] = row_return



if __name__ == '__main__':
    app = ScriptReplyTabel()
    from handlings.data_py import d_f
    datas  = app.main(r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\new\26,01,01-12,31.xlsx",d_f)
