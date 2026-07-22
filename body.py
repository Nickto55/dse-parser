import os
import sys
import plyer
import threading
import pandas as pd
import customtkinter as ctk

from PIL import Image
from tkinter import filedialog, END
from CTkMessagebox import CTkMessagebox

from dse_parcer import DseParser
from updater_result_table import UpdaterResultTableLogic
from scripts.excel_enter import ExcelDataInserter


def get_resource_path(relative_path):
    """ Возвращает абсолютный путь к ресурсу, учитывая сборку PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def resource_path(relative_path):
    try:
        from PIL import Image
        source_png = "static/png/parser_product.png"

        img = Image.open(source_png)

        icon_sizes = [(16, 16), (32, 32), (48, 48), (256, 256)]
        img.save("static/ico/app_icon.ico", sizes=icon_sizes)
        relative_path = "static/ico/app_icon.ico"
    except:
        print("Не удалось создать иконку")
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.normpath(os.path.join(base_path, relative_path))


path_to_main_ico = resource_path(r'static/ico/parser_product.ico')


def send_notification(title, message, name_program, settime=15):
    plyer.notification.notify(title=title, message=message, app_name=name_program, timeout=settime,
                              app_icon=resource_path(path_to_main_ico))


class HelpWindow(ctk.CTkToplevel):
    def __init__(self, parent=None):
        super().__init__()

        # Если parent передан, можно привязать окно к нему (опционально)
        if parent:
            self.transient(parent)

        self.title("DataWave - Справка")
        self.geometry('1010x600')

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

        self.run_param = 'regenerate'

        self.geomitri_constants()
        self.color_constant()

        self.create_gui()

    def geomitri_constants(self):
        """
        Основные значения для отрисовки GUI
        """
        """Main Window"""
        self.main_window_width = 1124
        self.main_window_height = 400
        """Indent's"""
        self.indent_self = 15
        self.indent_frame = 5

        ## Size
        """Size frames"""
        self.height_sidebar_frame = self.main_window_height - 2 * self.indent_self
        self.width_sidebar_frame = 225
        self.height_main_frame = 150
        self.width_main_frame = self.main_window_width - self.width_sidebar_frame - 3 * self.indent_self
        self.height_logs_frame = self.main_window_height - self.height_main_frame - self.indent_self
        self.width_logs_frame = self.main_window_width - self.width_sidebar_frame - 3 * self.indent_self

        '''aggregate size'''
        self.height_row_in_frame = 30
        self.width_button_in_frame = 100

        '''Sidebar sizes'''
        self.height_icon_frame = 3 * self.height_row_in_frame + 2 * self.indent_frame
        self.width_icon_frame = self.width_sidebar_frame - 2 * self.indent_frame
        self.height_logo_sidebar_frame = self.height_icon_frame - 2 * self.indent_frame
        self.width_logo_sidebar_frame = self.height_icon_frame - 2 * self.indent_frame
        self.height_checkbox_sidebar_frame = 3 * self.height_row_in_frame + 4 * self.indent_frame
        self.width_checkbox_sidebar_frame = self.width_sidebar_frame - 2 * self.indent_frame

        '''Main frame sizes'''
        self.width_path_entry = 480
        self.width_name_entry = 149

        self.position_gui()

    def color_constant(self):
        self.frame_color = '#434343'
        self.border_color = '#343638'

        self.fg_color_log_entry = '#1d1e1e'

        self.mode_selection_button_fgcolor = '#323a37'
        self.mode_selection_button_hover_color = '#183a2d'

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
        self.name_program = "Product parser"
        self.title(self.name_program)
        self.geometry(f"{self.main_window_width}x{self.main_window_height}")
        ctk.set_appearance_mode("dark")

        self._sidebar_frame()
        self._main_frame()
        self._log_fame()

        self.update_mode_selection(self.run_param)

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

        icon_frame = ctk.CTkFrame(
            sidebar_frame
            , width=self.width_icon_frame
            , height=self.height_icon_frame
            , fg_color=self.border_color
        )
        icon_frame.place(
            x=self.indent_frame
            , y=self.indent_frame
        )
        icon_englie = ctk.CTkImage(
            light_image=Image.open(get_resource_path("static/png/parser_product.png"))
            , dark_image=Image.open(get_resource_path("static/png/parser_product.png"))
            , size=(self.height_logo_sidebar_frame, self.width_logo_sidebar_frame)
        )
        ctk.CTkButton(
            icon_frame
            , image=icon_englie
            , compound='top'
            , width=self.width_logo_sidebar_frame
            , height=self.height_logo_sidebar_frame
            , hover=False
            , fg_color=self.border_color
        ).place(
            x=(self.width_sidebar_frame - self.width_logo_sidebar_frame) / 2 - self.indent_frame - 9
            # Смещение из за кривого лого
            , y=1  # Смещение из за кривого лого
        )

        mode_selection_button_frame = ctk.CTkFrame(
            sidebar_frame
            , width=self.width_checkbox_sidebar_frame
            , height=self.height_checkbox_sidebar_frame
            , fg_color=self.fg_color_log_entry
        )

        mode_selection_button_frame.place(
            x=self.indent_frame
            , y=self.height_icon_frame + 2 * self.indent_frame
        )
        select_entry = ctk.CTkEntry(
            mode_selection_button_frame
            , width=self.indent_frame + 2 * self.width_button_in_frame
            , height=self.height_row_in_frame
            , corner_radius=4
            , state='readonly'
            , border_color=self.border_color
        )
        select_entry.place(
            x=self.indent_frame
            , y=self.indent_frame
        )
        select_entry.configure(text_color='#9aa5aa', state='normal')
        select_entry.delete(0, END)
        select_entry.insert(0, 'Выберите режим')
        select_entry.configure(state='readonly')

        self.button_select_auto = ctk.CTkButton(
            mode_selection_button_frame
            , width=self.indent_frame + 2 * self.width_button_in_frame
            , height=self.height_row_in_frame
            , fg_color=self.mode_selection_button_fgcolor
            , hover_color=self.mode_selection_button_hover_color
            , text='Автоматический'
            , command=lambda: self.update_mode_selection('auto')
        )
        self.button_select_auto.place(
            x=self.indent_frame
            , y=self.height_row_in_frame + 2 * self.indent_frame
        )
        self.button_select_regenerate = ctk.CTkButton(
            mode_selection_button_frame
            , width=self.width_button_in_frame
            , height=self.height_row_in_frame
            , fg_color=self.mode_selection_button_fgcolor
            , hover_color=self.mode_selection_button_hover_color
            , text='Пересоздать'
            , command=lambda: self.update_mode_selection('regenerate')
        )
        self.button_select_regenerate.place(
            x=self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )
        self.button_select_renovation = ctk.CTkButton(
            mode_selection_button_frame
            , width=self.width_button_in_frame
            , height=self.height_row_in_frame
            , fg_color=self.mode_selection_button_fgcolor
            , hover_color=self.mode_selection_button_hover_color
            , text='Обновить'
            , command=lambda: self.update_mode_selection('renovation')

        )
        self.button_select_renovation.place(
            x=self.width_button_in_frame + 2 * self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )
        self.button_select_auto.configure(fg_color=self.mode_selection_button_hover_color)



        art_frame = ctk.CTkFrame(
            sidebar_frame
            , width=self.width_checkbox_sidebar_frame
            , height=self.height_checkbox_sidebar_frame
            , fg_color=self.frame_color
            # , fg_color='red'
        )
        art_frame.place(
            x=self.indent_frame
            , y=self.height_checkbox_sidebar_frame + self.height_icon_frame + 2 * self.indent_frame
        )
        self.paint_art_frame(art_frame, self.width_checkbox_sidebar_frame, self.height_checkbox_sidebar_frame)

        help_button = ctk.CTkButton(
            sidebar_frame
            , text="help"
            , fg_color=self.border_color
            , hover_color=self.mode_selection_button_hover_color
            , width=self.width_checkbox_sidebar_frame
            , height=self.height_row_in_frame
            , command=lambda: HelpWindow(self)
        )
        help_button.place(
            x=self.indent_frame
            , y=self.height_sidebar_frame - self.height_row_in_frame - self.indent_frame
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
            , width=self.width_button_in_frame
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
            , width=self.width_button_in_frame
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
            , width=self.width_button_in_frame
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
            , width=self.width_button_in_frame
            , height=self.height_row_in_frame
            , command=self.run_manager_thread
        )
        self.start_button.place(
            x=self.width_name_entry + self.width_path_entry + self.width_button_in_frame + 4 * self.indent_frame
            , y=2 * self.height_row_in_frame + 3 * self.indent_frame
        )

        self.button_open_result_tabel = ctk.CTkButton(
            main_frame
            , width=self.width_button_in_frame
            , height=2 * self.height_row_in_frame + self.indent_frame + 1
            , text="Открыть\n результат"
            , command=self.command_button_open_result
            , fg_color='#b69765'
            , hover_color='#8f764f'
        )
        self.button_open_result_tabel.place(
            x=self.width_name_entry + self.width_path_entry + self.width_button_in_frame + 4 * self.indent_frame
            , y=self.indent_frame
        )
        self.button_open_result_tabel.place_forget()

    def _log_fame(self):
        self.height_logs_frame += 10
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
            , fg_color=self.fg_color_log_entry
        )
        self.status_text.place(x=self.indent_frame, y=self.indent_frame)

    def log(self, message):
        """Вывод логов в текстовое поле GUI"""
        self.status_text.insert("end", f"> {message}\n")
        self.status_text.see("end")

    def run_manager_thread(self):
        self.start_button.configure(state="disabled")
        if (
                self.reply_path_entry.get() == ''
                or self.detail_path_entry.get() == ''
                or self.product_path_entry.get() == ''
        ):
            if self.product_path_entry.get() == '':
                self.product_path_entry.configure(border_color="red")
                self.product_name_entry.configure(border_color="red")

            if self.reply_path_entry.get() == '':
                self.reply_path_entry.configure(border_color="red")
                self.reply_name_entry.configure(border_color="red")

            if self.detail_path_entry.get() == '':
                self.detail_path_entry.configure(border_color="red")
                self.detail_name_entry.configure(border_color="red")
            self.start_button.configure(state="normal")
            return

        self.log("Запуск основного класса...")
        thread = threading.Thread(target=self.execute_logic, daemon=True)
        thread.start()

    def button_path_commands(self, label_batton: str):
        if label_batton == 'product':
            try:
                data_input_path = list(self.open_fils_to_path(name='изделия'))
                if len(data_input_path) == 0:
                    data_input_path.append('')
                path_list_filr = data_input_path
            except:
                path_list_filr = ['']

            if path_list_filr[0] != '':
                str_paths = ""
                for path in path_list_filr: str_paths += f"{path}, "
                str_paths = str_paths[:-2]

                self.product_path_entry.delete(0, END)
                self.product_path_entry.insert(0, str_paths)

                self.product_name_entry.configure(text_color='#fff', state='normal')
                self.product_name_entry.delete(0, END)
                self.product_name_entry.insert(0, os.path.basename(str_paths))
                self.product_name_entry.configure(state='readonly')

                self.log(f"Установлен путь для файла изделия")

                self.product_path_entry.configure(border_color=self.border_color)
                self.product_name_entry.configure(border_color=self.border_color)

        if label_batton == 'detail':
            try:
                data_input_path = list(self.open_fils_to_path(name='детали'))
                if len(data_input_path) == 0:
                    data_input_path.append('')
                path_list_filr = data_input_path
            except:
                path_list_filr = ['']

            if path_list_filr[0] != '':
                str_paths = ""
                for path in path_list_filr: str_paths += f"{path}, "
                str_paths = str_paths[:-2]
                print(len(path_list_filr))

                self.detail_path_entry.delete(0, END)
                self.detail_path_entry.insert(0, str_paths)

                self.detail_name_entry.configure(text_color='#fff', state='normal')
                self.detail_name_entry.delete(0, END)
                if len(path_list_filr) > 1:
                    self.detail_name_entry.insert(0, f'Файлов: {len(path_list_filr)}')
                else:
                    self.detail_name_entry.insert(0, os.path.basename(str_paths))
                self.detail_name_entry.configure(state='readonly')

                self.log(f"Установлен путь для файла детали")

                self.detail_path_entry.configure(border_color=self.border_color)
                self.detail_name_entry.configure(border_color=self.border_color)

        if label_batton == 'reply':
            try:
                data_input_path = list(self.open_fils_to_path(name='отчетов'))
                if len(data_input_path) == 0:
                    data_input_path.append('')
                path_list_filr = data_input_path
            except:
                path_list_filr = ['']

            if path_list_filr[0] != '':
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

                self.reply_name_entry.configure(text_color='#fff', state='normal')
                self.reply_name_entry.delete(0, END)
                if len(path_list_filr) > 1:
                    self.reply_name_entry.insert(0, f'Файлов: {len(path_list_filr)}')
                else:
                    self.reply_name_entry.insert(0, os.path.basename(str_paths))
                self.reply_name_entry.configure(state='readonly')

                self.log(f"Установлен путь для файла отчета")

                self.reply_path_entry.configure(border_color=self.border_color)
                self.reply_name_entry.configure(border_color=self.border_color)

    def open_fils_to_path(self, name):
        try:
            filepaths = filedialog.askopenfilenames(
                title=f"Выберите Excel файлы для {name}",
                filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
            )
            if not filepaths:
                return ''
            return filepaths
        except Exception as erro:
            self.log('Error open tk folder:')
            self.log(erro)
            return ''

    def execute_logic(self):
        # try:

        def run_dse_parser():
            parcser_product = DseParser()
            data_result = parcser_product.main(
                self.product_path_entry.get().replace(", ", ",").split(",")
                , self.detail_path_entry.get().replace(", ", ",").split(",")
                , self.reply_path_entry.get().replace(", ", ",").split(",")
            )

            self.path_outfile = self.product_path_entry.get().replace(", ", ",").split(",")[0]

            inserter = ExcelDataInserter(self.path_outfile)
            inserter.insert_data(data_result, sheet_name="Изделия")
            inserter.close()

        def run_update_result_table():
            self.path_outfile = self.product_path_entry.get()

            updater_result_tabel = UpdaterResultTableLogic(
                file_path_main=self.path_outfile
                , file_path_cz=self.detail_path_entry.get().replace(", ", ",").split(",")
            )
            data_result, headers = updater_result_tabel.main()

            inserter = ExcelDataInserter(self.path_outfile)
            inserter.insert_data({'':data_result}, sheet_name="Изделия", headers=headers)
            inserter.close()

        # self.run_param = 'regenerate'
        if self.run_param == 'regenerate':
            run_dse_parser()
        if self.run_param == 'renovation':
            run_update_result_table()

        self.button_open_result_tabel.place(
            x=self.width_name_entry + self.width_path_entry + self.width_button_in_frame + 4 * self.indent_frame
            , y=self.indent_frame
        )

        self.log("Процесс успешно завершен.")
        # нужно добавить уведомление

        self.start_button.configure(state="normal")

    def paint_art_frame(self, art_frame, width, height):
        pass

    def command_button_open_result(self):
        def merge_color():
            self.button_open_result_tabel.configure(fg_color='#8f764f', hover_color='#5c4b32')

        self.button_open_result_tabel.after(1000, merge_color)

        try:
            os.startfile(self.path_outfile)
            self.log("-Файл открыт")
        except Exception as e:
            self.log(f"Ошибка при открытии файла: {e}")
            self.start_button.configure(state="normal")
            return

        try:
            send_notification(
                f"Файл открыт: {os.path.basename(self.path_outfile).replace('.xlsx', '')}"
                , ""
                , self.name_program
                , 16
            )
        except:
            send_notification(
                f"Файл открыт: {os.path.basename(self.path_outfile)}"
                , ""
                , self.name_program
                , 16
            )
        self.button_open_result_tabel.after(5000, self.start_button.configure(state="normal"))

    def update_mode_selection(self, params_start):
        if params_start == 'auto':
            # self.button_select_auto.configure(
            #     fg_color=self.mode_selection_button_hover_color, hover_color=self.mode_selection_button_hover_color)
            # self.button_select_regenerate.configure(
            #     fg_color=self.mode_selection_button_fgcolor, hover_color=self.mode_selection_button_hover_color)
            # self.button_select_renovation.configure(
            #     fg_color=self.mode_selection_button_fgcolor, hover_color=self.mode_selection_button_hover_color)
            # self.run_param = 'auto'
            self.log(f'Пока не реализовано')
        if params_start == 'regenerate':
            self.button_select_auto.configure(
                fg_color=self.mode_selection_button_fgcolor, hover_color=self.mode_selection_button_hover_color)
            self.button_select_regenerate.configure(
                fg_color=self.mode_selection_button_hover_color, hover_color=self.mode_selection_button_hover_color)
            self.button_select_renovation.configure(
                fg_color=self.mode_selection_button_fgcolor, hover_color=self.mode_selection_button_hover_color)
            self.run_param = 'regenerate'
        if params_start == 'renovation':
            self.button_select_auto.configure(
                fg_color=self.mode_selection_button_fgcolor, hover_color=self.mode_selection_button_hover_color)
            self.button_select_regenerate.configure(
                fg_color=self.mode_selection_button_fgcolor, hover_color=self.mode_selection_button_hover_color)
            self.button_select_renovation.configure(
                fg_color=self.mode_selection_button_hover_color, hover_color=self.mode_selection_button_hover_color)
            self.run_param = 'renovation'

        self.update_gui_for_mode_selection()
        print(params_start)

    def update_gui_for_mode_selection(self):
        if self.run_param == 'regenerate':
            self.product_path_entry.configure(placeholder_text='Введите путь к файлу/файлам изделий.')
            self.detail_path_entry.configure(placeholder_text='Введите путь к файлу/файлам деталей.')

            self.reply_path_entry.delete(0, END)
            self.reply_path_entry._activate_placeholder()

            self.reply_path_entry.place(
                x=self.width_name_entry + 2 * self.indent_frame
                , y=2 * self.height_row_in_frame + 3 * self.indent_frame
            )
            self.reply_name_entry.place(
                x=self.indent_frame
                , y=2 * self.height_row_in_frame + 3 * self.indent_frame
            )
            self.button_open_folder_reply.place(
                x=self.width_name_entry + self.width_path_entry + 3 * self.indent_frame
                , y=2 * self.height_row_in_frame + 3 * self.indent_frame
            )
            self.reply_name_entry.configure(text_color='#9aa5aa', state='normal')
            self.reply_name_entry.delete(0, END)
            self.reply_name_entry.insert(0, 'Имя файла')
            self.reply_name_entry.configure(state='readonly')


        if self.run_param == "renovation":

            self.product_path_entry.configure(placeholder_text='Введите путь к файлу Переводы по иделиям, для обновления данных.')
            self.detail_path_entry.configure(placeholder_text='Введите путь к файлу ДСЕ по СЗ.')

            self.reply_path_entry.delete(0, END)
            self.reply_path_entry.insert(0, 'None')
            self.reply_name_entry.configure(text_color='#fff', state='normal')
            self.reply_name_entry.delete(0, END)
            self.reply_name_entry.insert(0, os.path.basename('None'))
            self.reply_name_entry.configure(state='readonly')
            self.reply_path_entry.place_forget()
            self.reply_name_entry.place_forget()
            self.button_open_folder_reply.place_forget()




if __name__ == "__main__":
    app = AppGui()
    app.mainloop()
