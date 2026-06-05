from scripts.handlings.handler_main_tabel import HandlerProductTabel
from scripts.handlings.handler_detail_tabel import HandlerDetailTabel
from scripts.reply_script import ScriptReplyTabel

class DseParser:
    def __init__(self):

        self.handler_pTabel = HandlerProductTabel()
        self.handler_dTabel = HandlerDetailTabel()

        self.scr_reply = ScriptReplyTabel()

        self.data_product_tabel = {}
        self.data_detail_tabel = {}

        self.data = {}

    def main(self, list_paths_product_tabels: list, list_paths_detail_tabels: list, list_paths_reply_tabels: list):
        for path_to_detail_taibel in list_paths_detail_tabels:
            self.data_detail_tabel = self.handler_dTabel.main(path_to_detail_tabel=path_to_detail_taibel,
                                                              data=self.data_detail_tabel)

        count_dse = 0
        for path_to_product_tabel in list_paths_product_tabels:
            self.data_product_tabel = self.handler_pTabel.main(path_to_product_tabel)
            key_list_product_data = list((self.data_product_tabel.keys()))

            for dse_detail in self.data_detail_tabel.keys():
                for dse_key in key_list_product_data:
                    if dse_key in dse_detail:
                        count_dse += 1
                        self.data[dse_detail] = {
                            dse_detail: {
                                'ДСЕ (детали)': dse_key
                                , 'Наименование (детали)': self.data_product_tabel[dse_key].get('name', '')
                                , 'ДСЕ (изделия)': self.data_detail_tabel[dse_detail].get('dse', '')
                                , 'Наименование (изделия)': self.data_detail_tabel[dse_detail].get('detai', '')
                                , 'частота упоминаний': self.data_detail_tabel[dse_detail].get('count dse', '')
                            }
                        }

        for path_to_reply_file in list_paths_reply_tabels:
            self.data = self.scr_reply.main(path_reply_tabl_file=path_to_reply_file, receive_data=self.data).copy()

        return self.data


if __name__ == '__main__':
    app = DseParser()
    data_product = app.main(
        list_paths_product_tabels=[r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\Копия Список переводов 28.05.xlsx"]
        , list_paths_detail_tabels=[
            r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\ер6.212.112.xls"
            , r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\ер7.756.090.xls"
            , r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\лтия.723116.033.xls"
        ]
    )

    for row_num, row in data_product.items():
        print(row_num)
        for key_num, keyd in row.items():
            print("     ", key_num, keyd)
