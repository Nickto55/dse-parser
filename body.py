import threading
import customtkinter as ctk
from tkinter import filedialog, END

from dse_parcer import DseParser
from scripts.excel_enter import ExcelDataInserter

class AppGui(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Product parser")
        self.geometry("530x420")
        ctk.set_appearance_mode("dark")

        self.main_frame = ctk.CTkFrame(
            self
            ,width=510
            ,fg_color="transparent"
        )
        self.main_frame.pack(pady=10, padx=10, fill="x")


        self.product_path_entry = ctk.CTkEntry(
            self.main_frame
            ,width=380
            ,height=30
            ,corner_radius=4
            ,placeholder_text='Введите путь к файлу/файлам изделий.'
        )
        self.product_path_entry.grid(row=0, column=0, padx=(0, 10), pady=10, sticky="w")

        self.button_open_folder_product = ctk.CTkButton(
            self.main_frame
            ,text='Открыть'
            ,width=100
            ,height=30
            ,command=lambda : self.button_path_commands(label_batton='product')
        )
        self.button_open_folder_product.grid(row=0, column=1, pady=10, sticky="e")


        self.detail_path_entry = ctk.CTkEntry(
            self.main_frame
            ,width=380
            ,height=30
            ,corner_radius=4
            ,placeholder_text='Введите путь к файлу/файлам деталей.'
        )
        self.detail_path_entry.grid(row=1, column=0, padx=(0, 10), pady=10, sticky="w")

        self.button_open_folder_detail = ctk.CTkButton(
            self.main_frame
            ,text='Открыть'
            ,width=100
            ,height=30
            ,command=lambda : self.button_path_commands(label_batton='detail')
        )
        self.button_open_folder_detail.grid(row=1, column=1, pady=10, sticky="e")


        self.status_text = ctk.CTkTextbox(self, width=510, height=180)
        self.status_text.pack(pady=5, padx=10)
        self.status_text.insert("0.0", "Готов к запуску...\n")

        self.start_button = ctk.CTkButton(
            self
            ,text="Начать"
            ,fg_color="green"
            ,hover_color="darkgreen"
            ,width=150
            ,height=35
            ,command=self.run_manager_thread
        )
        self.start_button.pack(pady=15)

    def log(self, message):
        """Вывод логов в текстовое поле GUI"""
        self.status_text.insert("end", f"> {message}\n")
        self.status_text.see("end")

    def run_manager_thread(self):
        """Запуск в отдельном потоке, чтобы GUI не зависал"""
        self.start_button.configure(state="disabled")
        self.log("Запуск основного класса...")
        
        thread = threading.Thread(target=self.execute_logic, daemon=True)
        thread.start()
    
    def button_path_commands(self, label_batton: str):
        if label_batton == 'product':
            path_list_filr = list(self.open_fils_to_path())

            str_paths = ""
            for path in path_list_filr: str_paths += f"{path}, "
            str_paths = str_paths[:-2]

            self.product_path_entry.delete(0, END)
            self.product_path_entry.insert(0, str_paths)
            self.log(f"Установлен путь для файла изделия")
        if label_batton == 'detail':
            path_list_filr = list(self.open_fils_to_path())

            str_paths = ""
            for path in path_list_filr: str_paths += f"{path}, "
            str_paths = str_paths[:-2]

            self.detail_path_entry.delete(0, END)
            self.detail_path_entry.insert(0, str_paths)
            self.log(f"Установлен путь для файла детали")
    
    def open_fils_to_path(self):
        filepaths = filedialog.askopenfilenames(
            title="Выберите Excel файлы",
            filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
        )
        if not filepaths:
            return
        return filepaths

    def execute_logic(self):
        try:
            parcser_product =  DseParser()
            data_result = parcser_product.main(
                self.product_path_entry.get().replace(", ", ",").split(",")
                ,self.detail_path_entry.get().replace(", ", ",").split(",")
            )

            # for row_num, row in data_result.items():
            #     print(row_num)
            #     for key_num,keyd in row.items():
            #         print("     ",key_num,keyd)

            inserter = ExcelDataInserter(self.product_path_entry.get().replace(", ", ",").split(",")[0])
            inserter.insert_data(data_result, sheet_name="Изделия")
            inserter.close()
            
            self.log("Процесс успешно завершен.")
        except Exception as e:
            self.log(f"ОШИБКА: {str(e)}")
        finally:
            self.start_button.configure(state="normal")


if __name__ == "__main__":
    app = AppGui()
    app.mainloop()
