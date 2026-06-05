import threading
import customtkinter as ctk

from tkinter import filedialog, END
from CTkMessagebox import CTkMessagebox

from dse_parcer import DseParser
from scripts.excel_enter import ExcelDataInserter


class HelpWindow(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__()

        # Если parent передан, можно привязать окно к нему (опционально)
        if parent:
            self.transient(parent)

        self.title("DataWave - Справка")
        self.geometry('700x600')

        # Создаем текстовое поле. В CTk нет поддержки HTML, только обычный текст.
        # state="normal" нужен для вставки текста, потом переключим на "disabled" (read-only)
        self.text_edit = ctk.CTkTextbox(self, state="normal", wrap="word")
        self.text_edit.pack(fill="both", expand=True, padx=20, pady=(20, 10))

        try:
            from static.help_text import help_text as help_text_str
            # Вставляем текст с начала (индекс "0.0")
            self.text_edit.insert("0.0", help_text_str)
            self.text_edit.configure(state="disabled")  # Делаем поле только для чтения

        except Exception as e:
            self.text_edit.insert("0.0", f"Ошибка загрузки справки:\n{str(e)}")
            self.text_edit.configure(state="disabled")

        # Кнопка закрытия (аналог QPushButton)
        # width=100 заменяет setFixedWidth(100)
        self.close_btn = ctk.CTkButton(
            self,
            text="Закрыть",
            width=100,
            command=self.destroy  # Аналог clicked.connect(self.close)
        )
        # pack с anchor="e" и side="bottom" заменяет QHBoxLayout с addStretch()
        # (кнопка будет прижата к правому нижнему углу)
        self.close_btn.pack(side="bottom", anchor="e", padx=20, pady=(0, 20))

    def show_event(self):
        """
        Альтернативный способ показа окна (аналог exec() в PyQt).
        Делает окно видимым, поднимает его наверх и передает фокус.
        """
        self.deiconify()  # Показать окно, если оно было скрыто
        self.lift()  # Поднять на передний план
        self.focus_force()  # Передать фокус

        # Если нужно модальное поведение (блокировка основного окна, как в exec()):
        # self.grab_set()
        # self.wait_window()



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



        self.reply_path_entry = ctk.CTkEntry(
            self.main_frame
            , width=380
            , height=30
            , corner_radius=4
            , placeholder_text='Введите путь к файлу/файлам отчетов'
        )
        self.reply_path_entry.grid(row=2, column=0, padx=(0, 10), pady=10, sticky="w")

        self.button_open_folder_reply = ctk.CTkButton(
            self.main_frame
            , text='Открыть'
            , width=100
            , height=30
            , command=lambda: self.button_path_commands(label_batton='reply')
        )
        self.button_open_folder_reply.grid(row=2, column=1, pady=10, sticky="e")




        self.status_text = ctk.CTkTextbox(self, width=510, height=180)
        self.status_text.pack(pady=5, padx=10)
        self.status_text.insert("0.0", "Готов к запуску...\n")

        help_button = ctk.CTkButton(
            self
            ,text="Help"
            ,fg_color='green'
            ,hover_color='darkgrey'
            ,width=50
            ,height=35
            ,command=lambda : HelpWindow(self)
        )


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
        help_button.pack(pady=15)

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
            path_list_filr = list(self.open_fils_to_path(name='изделия'))

            str_paths = ""
            for path in path_list_filr: str_paths += f"{path}, "
            str_paths = str_paths[:-2]

            self.product_path_entry.delete(0, END)
            self.product_path_entry.insert(0, str_paths)
            self.log(f"Установлен путь для файла изделия")
        if label_batton == 'detail':
            path_list_filr = list(self.open_fils_to_path(name='детали'))

            str_paths = ""
            for path in path_list_filr: str_paths += f"{path}, "
            str_paths = str_paths[:-2]

            self.detail_path_entry.delete(0, END)
            self.detail_path_entry.insert(0, str_paths)
            self.log(f"Установлен путь для файла детали")
        if label_batton == 'reply':
            path_list_filr = list(self.open_fils_to_path(name='отчетов'))

            for path_file in path_list_filr:
                if ',' in path_file:
                    CTkMessagebox(
                        title='Внимание!'
                        ,message='замените в названии файла или переименуйте местоположение файла так, что бы, в нем не было зяпятой ,'
                        , icon='warning'
                    )
                    return None
            str_paths = ""
            for path in path_list_filr: str_paths += f"{path}, "
            str_paths = str_paths[:-2]

            self.reply_path_entry.delete(0, END)
            self.reply_path_entry.insert(0, str_paths)
            self.log(f"Установлен путь для файла jnxtnjd")
    
    def open_fils_to_path(self, name):
        filepaths = filedialog.askopenfilenames(
            title=f"Выберите Excel файлы для ХЪ{name}",
            filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
        )
        if not filepaths:
            return
        return filepaths

    def execute_logic(self):
        # try:
        parcser_product =  DseParser()
        data_result = parcser_product.main(
            self.product_path_entry.get().replace(", ", ",").split(",")
            ,self.detail_path_entry.get().replace(", ", ",").split(",")
            ,self.reply_path_entry.get().replace(", ", ",").split(",")
        )


        inserter = ExcelDataInserter(self.product_path_entry.get().replace(", ", ",").split(",")[0])
        inserter.insert_data(data_result, sheet_name="Изделия")
        inserter.close()

        self.log("Процесс успешно завершен.")
        # except Exception as e:
        #     self.log(f"ОШИБКА: {str(e)}")
        # finally:
        #     self.start_button.configure(state="normal")


if __name__ == "__main__":
    app = AppGui()
    app.mainloop()
