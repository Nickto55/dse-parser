
from scripts.handlings.handler_main_tabel import HandlerProductTabel
from scripts.handlings.handler_detail_tabel import HandlerDetailTabel

class DseParser:
    def __init__(self):
        
        self.handler_pTabel = HandlerProductTabel()
        self.handler_dTabel = HandlerDetailTabel()

        self.data_product_tabel = {}
        self.data_detail_tabel = {}

        self.data = {}

    def main(self, list_paths_product_tabels: list, list_paths_detail_tabels: list):
        for path_to_detail_taibel in list_paths_detail_tabels:
            self.data_detail_tabel = self.handler_dTabel.main(path_to_detail_tabel=path_to_detail_taibel, data=self.data_detail_tabel)

        for path_to_product_tabel in list_paths_product_tabels:
            self.data_product_tabel = self.handler_pTabel.main(path_to_product_tabel)
            key_list_product_data = list((self.data_product_tabel.keys()))
            
            for dse_detail in self.data_detail_tabel.keys():
                if not '_+_' in dse_detail: continue
                part = dse_detail[:dse_detail.index('_+_')]
                if part.lower() in key_list_product_data:
                    self.data[part.lower()]={
                        dse_detail[dse_detail.index('_+_')+3:]:{
                            'ДСЕ (изделия)' : part.upper()
                            ,'ДСЕ (детали)':self.data_detail_tabel[dse_detail].get('dse','')
                            ,'Наименование (детали)':self.data_detail_tabel[dse_detail].get('detai','')
                        }
                    }
        return self.data
                    

            
                
        

        
            



if __name__ == '__main__':
    app = DseParser()
    data_product = app.main(
        list_paths_product_tabels=[r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\Копия Список переводов 28.05.xlsx"]
        ,list_paths_detail_tabels=[
            r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\ер6.212.112.xls"
            ,r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\ер7.756.090.xls"
            ,r"C:\Users\yakovlev_nd\Desktop\Tests\DseParser\отчеты\лтия.723116.033.xls"
        ]
    )

    for row_num, row in data_product.items():
        print(row_num)
        for key_num,keyd in row.items():
            print("     ",key_num,keyd)

