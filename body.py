import threading
import customtkinter as ctk

from tkinter import filedialog, END

import pandas as pd
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
        self.geomitri_constants()
        self.color_constant()

        self.create_gui()

    def geomitri_constants(self):
        """
        Основные значения для отрисовки GUI
        """
        """Main Window"""
        self.main_window_width = 1119
        self.main_window_height = 640
        """Indent's"""
        self.indent_self = 15
        self.indent_frame = 5

        ## Size
        """Size frames"""
        self.height_sidebar_frame = self.main_window_height - 2 * self.indent_self
        self.width_sidebar_frame = 220

        self.height_main_frame = 150
        self.width_main_frame = self.main_window_width - self.width_sidebar_frame - 3 * self.indent_self

        self.height_logs_frame = self.height_sidebar_frame - self.height_main_frame - self.indent_self
        self.width_logs_frame = self.main_window_width - self.width_sidebar_frame - 3 * self.indent_self

        '''aggregate size'''
        self.height_row_in_frame = 30

        self.width_botton_in_frame = 100

        '''Main frame sizes'''
        self.width_path_entry = 480
        self.width_name_entry = 149

        self.position_gui()

    def color_constant(self):
        self.frame_color = '#434343'
        self.border_color = '#343638'

        """button"""
        self.button_fg_color = '#475563'
        self.button_hover_color = '#353c43'

    def position_gui(self):
        """position frames"""
        self.pos_sidebar_frame_x = self.indent_self
        self.pos_sidebar_frame_y = self.indent_self

        self.pos_main_frame_x = self.pos_sidebar_frame_x + self.width_sidebar_frame + self.indent_self
        self.pos_main_frame_y = self.pos_sidebar_frame_y

        self.pos_logs_frame_x = self.pos_sidebar_frame_x + self.width_sidebar_frame + self.indent_self
        self.pos_logs_frame_y = self.pos_main_frame_y + self.height_main_frame + self.indent_self

    def create_gui(self):
        self.title("Product parser")
        self.geometry(f"{self.main_window_width}x{self.main_window_height}")
        ctk.set_appearance_mode("dark")

        self._sidebar_frame()
        self._main_frame()
        self._log_fame()

        # self.start_button = ctk.CTkButton(
        #     self
        #     ,text="Начать"
        #     ,fg_color="green"
        #     ,hover_color="darkgreen"
        #     ,width=150
        #     ,height=35
        #     ,command=self.run_manager_thread
        # )
        # self.start_button.pack(pady=15)
        # help_button.pack(pady=15)

    def _sidebar_frame(self):
        sidebar_frame = ctk.CTkFrame(
            self
            , width=self.width_sidebar_frame
            , height=self.height_sidebar_frame
            , fg_color=self.frame_color
        )
        sidebar_frame.place(
            x=self.pos_sidebar_frame_y
            , y=self.pos_sidebar_frame_y
        )

    def _main_frame(self):
        self.height_main_frame = 3 * self.height_row_in_frame + 4 * self.indent_frame

        main_frame = ctk.CTkFrame(
            self
            , width=self.width_main_frame
            , height=self.height_main_frame
            , fg_color=self.frame_color
        )
        main_frame.place(
            x=self.pos_main_frame_x
            , y=self.pos_main_frame_y
        )

        ### Ввод данных для списка переводов ###
        self.product_name_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_name_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Имя файла'
            , state='readonly'
            , border_color=self.border_color
        )
        self.product_name_entry.place(
            x=self.indent_frame
            , y=self.indent_frame
        )
        self.product_name_entry.configure(text_color='#9aa5aa', state='normal')
        self.product_name_entry.delete(0, END)
        self.product_name_entry.insert(0, 'Имя файла')
        self.product_name_entry.configure(state='readonly')

        self.product_path_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_path_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Введите путь к файлу/файлам изделий.'
            , border_color=self.border_color
        )
        self.product_path_entry.place(
            x=self.width_name_entry + 2 * self.indent_frame
            , y=self.indent_frame
        )

        self.button_open_folder_product = ctk.CTkButton(
            main_frame
            , text='Открыть'
            , width=self.width_botton_in_frame
            , height=self.height_row_in_frame
            , command=lambda: self.button_path_commands(label_batton='product')
            , fg_color=self.button_fg_color
            , hover_color=self.button_hover_color
        )
        self.button_open_folder_product.place(
            x=self.width_name_entry + self.width_path_entry + 3 * self.indent_frame
            , y=self.indent_frame
        )

        ### Ввод данных для списка очетов ###

        self.detail_name_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_name_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Имя файла'
            , state='readonly'
            , border_color=self.border_color
        )
        self.detail_name_entry.place(
            x=self.indent_frame
            , y=self.height_row_in_frame + 2 * self.indent_frame
        )
        self.detail_name_entry.configure(text_color='#9aa5aa', state='normal')
        self.detail_name_entry.delete(0, END)
        self.detail_name_entry.insert(0, 'Имя файла')
        self.detail_name_entry.configure(state='readonly')

        self.detail_path_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_path_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Введите путь к файлу/файлам деталей.'
            , border_color=self.border_color
        )
        self.detail_path_entry.place(
            x=self.width_name_entry + 2 * self.indent_frame
            , y=self.height_row_in_frame + 2 * self.indent_frame
        )

        self.button_open_folder_detail = ctk.CTkButton(
            main_frame
            , text='Открыть'
            , width=self.width_botton_in_frame
            , height=self.height_row_in_frame
            , command=lambda: self.button_path_commands(label_batton='detail')
            , fg_color=self.button_fg_color
            , hover_color=self.button_hover_color
        )
        self.button_open_folder_detail.place(
            x=self.width_name_entry + self.width_path_entry + 3 * self.indent_frame
            , y=self.height_row_in_frame + 2 * self.indent_frame
        )

        ### Ввод данных для годового отчета ###

        self.reply_name_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_name_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Имя файла'
            , state='readonly'
            , border_color=self.border_color
        )
        self.reply_name_entry.place(
            x=self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )
        self.reply_name_entry.configure(text_color='#9aa5aa', state='normal')
        self.reply_name_entry.delete(0, END)
        self.reply_name_entry.insert(0, 'Имя файла')
        self.reply_name_entry.configure(state='readonly')

        self.reply_path_entry = ctk.CTkEntry(
            main_frame
            , width=self.width_path_entry
            , height=self.height_row_in_frame
            , corner_radius=4
            , placeholder_text='Введите путь к файлу/файлам отчетов'
            , border_color=self.border_color
        )
        self.reply_path_entry.place(
            x=self.width_name_entry + 2 * self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )

        self.button_open_folder_reply = ctk.CTkButton(
            main_frame
            , text='Открыть'
            , width=self.width_botton_in_frame
            , height=self.height_row_in_frame
            , command=lambda: self.button_path_commands(label_batton='reply')
            , fg_color=self.button_fg_color
            , hover_color=self.button_hover_color
        )
        self.button_open_folder_reply.place(
            x=self.width_name_entry + self.width_path_entry + 3 * self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )

        self.start_button = ctk.CTkButton(
            main_frame
            , text="Начать"
            , fg_color='#305433'
            , hover_color='#2d432e'
            , width=self.width_botton_in_frame
            , height=self.height_row_in_frame
            , command=self.run_manager_thread
        )
        self.start_button.place(
            x=self.width_name_entry + self.width_path_entry + self.width_botton_in_frame + 4 * self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )

    def _log_fame(self):
        self.position_gui()
        logs_frame = ctk.CTkFrame(
            self
            , width=self.width_logs_frame
            , height=self.height_logs_frame
            , fg_color=self.frame_color
        )
        logs_frame.place(x=self.pos_logs_frame_x, y=self.pos_logs_frame_y)

        self.status_text = ctk.CTkTextbox(
            logs_frame
            , width=self.width_logs_frame - 2 * self.indent_frame
            , height=self.height_logs_frame - 2 * self.indent_frame
        )
        self.status_text.place(x=self.indent_frame, y=self.indent_frame)

    def log(self, message):
        """Вывод логов в текстовое поле GUI"""
        self.status_text.insert("end", f"> {message}\n")
        self.status_text.see("end")

    def run_manager_thread(self):
        """Запуск в отдельном потоке, чтобы GUI не зависал"""
        self.start_button.configure(state="disabled")
        if not pd.isna(self.reply_path_entry.get()) or not pd.isna(self.detail_path_entry.get()) or not pd.isna(
                self.product_path_entry.get()):
            self.log("Запуск dksoapsdkoap;d")
            return
        print(pd.isna(self.reply_path_entry.get()) or pd.isna(self.detail_path_entry.get()) or pd.isna(
            self.product_path_entry.get()))
        print(pd.isna(self.reply_path_entry.get()), pd.isna(self.detail_path_entry.get()),
              pd.isna(self.product_path_entry.get()))
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
                        ,
                        message='замените в названии файла или переименуйте местоположение файла так, что бы, в нем не было зяпятой ,'
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
            title=f"Выберите Excel файлы для {name}",
            filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
        )
        if not filepaths:
            return
        return filepaths

    def execute_logic(self):
        # try:
        parcser_product = DseParser()
        data_result = parcser_product.main(
            self.product_path_entry.get().replace(", ", ",").split(",")
            , self.detail_path_entry.get().replace(", ", ",").split(",")
            , self.reply_path_entry.get().replace(", ", ",").split(",")
        )

        inserter = ExcelDataInserter(self.product_path_entry.get().replace(", ", ",").split(",")[0])
        inserter.insert_data(data_result, sheet_name="Изделия")
        inserter.close()

        self.log("Процесс успешно завершен.")


if __name__ == "__main__":
    app = AppGui()
    app.mainloop()
