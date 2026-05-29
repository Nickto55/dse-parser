from tkinter import messagebox
from typing import Optional, Union, List, Dict, Any, Hashable

import pandas as pd
import openpyxl
from pathlib import Path

try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False


class ExcelReader:
    """
    Базовый класс для чтения Excel-файлов.
    Автоматически определяет формат (.xls / .xlsx) и использует правильный парсер.
    Для старых .xls файлов пробует xlrd, затем pandas с engine='xlrd'.
    """

    # Популярные русские кодировки для старых Excel-файлов
    ENCODINGS = ["cp1251", "utf-8", "cp866", "koi8_r", "iso-8859-5"]

    def __init__(self, file_path: str, sheet_name: Optional[Union[str, int]] = None,
                 color_filter_column: Optional[str] = None,
                 track_sheet_origin: bool = False,
                 encoding: Optional[str] = None):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.encoding = encoding
        self.data: Optional[pd.DataFrame] = None
        self.columns_save: List[str] = []
        self.filtered_data: Optional[Dict[int, Dict[str, Any]]] = None
        self.color_filter_column = color_filter_column
        self.track_sheet_origin = track_sheet_origin
        self.sheet_origin: Optional[str] = None

        self.load_excel()

    @staticmethod
    def fix_encoding(text: Any) -> Any:
        """
        Исправляет типичные проблемы с кодировкой (cp1252->cp1251).
        Если текст уже корректный — возвращает как есть.
        """
        if not isinstance(text, str):
            return text

        try:
            return text.encode('cp1252', errors='ignore').decode('cp1251', errors='ignore')
        except (UnicodeEncodeError, UnicodeDecodeError):
            return text

    @staticmethod
    def fix_dataframe_encoding(df: pd.DataFrame) -> pd.DataFrame:
        """
        Рекурсивно исправляет кодировку во всех строковых значениях DataFrame.
        """
        df_fixed = df.copy()
        for col in df_fixed.columns:
            if df_fixed[col].dtype == object:
                df_fixed[col] = df_fixed[col].apply(ExcelReader.fix_encoding)
        return df_fixed

    def _detect_format(self, file_path: str) -> str:
        """Определяет формат файла по расширению или бинарной сигнатуре."""
        path = Path(file_path)
        ext = path.suffix.lower()
        if ext == ".xls":
            return "xls"
        elif ext in (".xlsx", ".xlsm"):
            return "xlsx"
        else:
            with open(file_path, "rb") as f:
                header = f.read(8)
            if header[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                return "xls"
            elif header[:4] == b'PK\x03\x04':
                return "xlsx"
            else:
                raise ValueError(f"Неизвестный формат файла: {file_path}")

    def load_excel(self, file_path: Optional[str] = None,
                   sheet_name: Optional[Union[str, int]] = None,
                   color_filter_column: Optional[str] = None,
                   encoding: Optional[str] = None) -> Optional[pd.DataFrame]:
        """
        Загружает данные из Excel файла с автоматическим определением формата
        и исправлением кодировки.
        """
        path = file_path or self.file_path
        sheet = sheet_name if sheet_name is not None else self.sheet_name
        color_col = color_filter_column or self.color_filter_column
        enc = encoding or self.encoding

        file_format = self._detect_format(path)
        print(f"[ExcelReader] Обнаружен формат: {file_format}")

        try:
            if file_format == "xls":
                data = self._load_xls(path, sheet, color_col, enc)
            else:
                data = self._load_xlsx(path, sheet, color_col)

            if isinstance(data, pd.DataFrame):
                data = self.fix_dataframe_encoding(data)
            elif isinstance(data, dict):
                for key in data:
                    if isinstance(data[key], pd.DataFrame):
                        data[key] = self.fix_dataframe_encoding(data[key])

                sheet_names = list(data.keys())
                if sheet_names:
                    first_sheet = sheet_names[0]
                    print(f"Лист не указан. Загружаем первый лист: {first_sheet}")
                    data = data[first_sheet]
                else:
                    raise ValueError("Файл Excel не содержит листов.")

            if file_path is None and sheet_name is None:
                self.data = data
                print(f"Лист успешно загружен: {sheet if sheet else 'первый лист'}")

            return data

        except Exception as e:
            print(f"Ошибка при загрузке файла или листа: {e}")
            if file_path is None and sheet_name is None:
                self.data = None
            return None

    def _load_xls(self, path: str, sheet: Optional[Union[str, int]],
                  color_col: Optional[str], encoding: Optional[str]) -> pd.DataFrame:
        """Загружает .xls: пробует xlrd напрямую, затем pandas с engine='xlrd'."""

        # Сначала пробуем xlrd напрямую для лучшего контроля кодировки
        if XLRD_AVAILABLE:
            workbook = self._try_xlrd(path, encoding)
            if workbook:
                return self._read_xlrd_workbook(workbook, sheet, color_col)

        # Fallback: pandas с engine='xlrd' (xlrd сам подберёт кодировку)
        print("[ExcelReader] Fallback на pandas + xlrd engine")
        try:
            df = pd.read_excel(path, sheet_name=sheet, engine="xlrd")
            return df
        except Exception as e:
            print(f"[ExcelReader] pandas+xlrd тоже не сработал: {e}")

        # Последний fallback: пробуем как .xlsx (на случай если файл на самом деле xlsx)
        try:
            print("[ExcelReader] Пробуем открыть как .xlsx...")
            df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
            return df
        except Exception as e:
            raise ValueError(f"Не удалось открыть файл .xls ни одним способом. Последняя ошибка: {e}")

    def _try_xlrd(self, path: str, encoding: Optional[str]) -> Optional[Any]:
        """Пробует открыть .xls через xlrd с разными кодировками."""
        if not XLRD_AVAILABLE:
            return None

        # Принудительная кодировка
        if encoding:
            try:
                wb = xlrd.open_workbook(path, encoding_override=encoding)
                print(f"[ExcelReader] xlrd открыл с кодировкой: {encoding}")
                return wb
            except Exception as e:
                print(f"[ExcelReader] Кодировка {encoding} не подошла: {e}")

        # Автоподбор кодировок
        for enc in self.ENCODINGS:
            try:
                wb = xlrd.open_workbook(path, encoding_override=enc)
                print(f"[ExcelReader] xlrd открыл с кодировкой: {enc}")
                return wb
            except Exception:
                continue

        # Без encoding_override
        try:
            wb = xlrd.open_workbook(path)
            print(f"[ExcelReader] xlrd открыл без принудительной кодировки")
            return wb
        except Exception as e:
            print(f"[ExcelReader] xlrd не смог открыть файл: {e}")
            return None

    def _read_xlrd_workbook(self, workbook, sheet: Optional[Union[str, int]],
                            color_col: Optional[str]) -> pd.DataFrame:
        """Читает данные из xlrd workbook в pandas DataFrame."""
        # Выбираем лист
        if sheet is None:
            sheet_idx = 0
            self.sheet_origin = workbook.sheet_names()[0]
        elif isinstance(sheet, int):
            sheet_idx = sheet
            self.sheet_origin = workbook.sheet_names()[sheet]
        else:
            sheet_idx = workbook.sheet_names().index(sheet)
            self.sheet_origin = sheet

        ws = workbook.sheet_by_index(sheet_idx)

        # Читаем заголовки
        header_row = 0
        headers = []
        for col in range(ws.ncols):
            val = ws.cell(header_row, col).value
            headers.append(str(val) if val is not None else f"col_{col}")

        # Фильтрация по цвету не поддерживается в xlrd
        if color_col is not None:
            print("[ExcelReader] Предупреждение: фильтрация по цвету для .xls не поддерживается")

        # Читаем данные
        rows = []
        for row_idx in range(header_row + 1, ws.nrows):
            row_data = {}
            for col_idx in range(ws.ncols):
                key = headers[col_idx] if col_idx < len(headers) else f"col_{col_idx}"
                cell = ws.cell(row_idx, col_idx)
                row_data[key] = self._convert_xls_value(cell, workbook)
            rows.append(row_data)

        df = pd.DataFrame(rows)

        if self.track_sheet_origin:
            df['__sheet_origin__'] = self.sheet_origin

        return df

    def _convert_xls_value(self, cell, workbook):
        """Конвертирует xlrd-ячейку в нативный Python-тип."""
        if cell.ctype == xlrd.XL_CELL_EMPTY or cell.ctype == xlrd.XL_CELL_BLANK:
            return None
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            if cell.value == int(cell.value):
                return int(cell.value)
            return cell.value
        elif cell.ctype == xlrd.XL_CELL_DATE:
            return xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode)
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)
        elif cell.ctype == xlrd.XL_CELL_ERROR:
            return None
        else:
            return self.fix_encoding(cell.value)

    def _load_xlsx(self, path: str, sheet: Optional[Union[str, int]],
                   color_col: Optional[str]) -> pd.DataFrame:
        """Загружает .xlsx через openpyxl/pandas."""
        if color_col is not None:
            return self._load_xlsx_with_color_filter(path, sheet, color_col)
        else:
            return pd.read_excel(path, sheet_name=sheet)

    def _load_xlsx_with_color_filter(self, path: str, sheet: Optional[Union[str, int]],
                                     color_col: str) -> pd.DataFrame:
        """Загружает данные с цветной фильтрацией и исправляет кодировку."""
        wb = openpyxl.load_workbook(path, data_only=True)

        if sheet is None:
            ws = wb.active
            self.sheet_origin = ws.title
        elif isinstance(sheet, int):
            ws = wb.worksheets[sheet]
            self.sheet_origin = ws.title
        else:
            ws = wb[sheet]
            self.sheet_origin = sheet

        header_row = 1
        headers = [ExcelReader.fix_encoding(cell.value) for cell in ws[header_row]]

        try:
            col_idx = headers.index(color_col) + 1
        except ValueError:
            wb.close()
            raise ValueError(f"Колонка '{color_col}' не найдена")

        colored_rows = []
        row_colors = []

        for row_idx in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)

            color = None
            has_color = False

            if cell.fill and cell.fill.fgColor:
                rgb = cell.fill.fgColor.rgb
                if rgb and rgb not in ('00000000', 'FFFFFFFF', None):
                    color = str(rgb)
                    has_color = True

            if has_color:
                colored_rows.append(row_idx - header_row - 1)
                row_colors.append(color)

        wb.close()

        df = pd.read_excel(path, sheet_name=sheet)
        df = self.fix_dataframe_encoding(df)
        filtered_df = df.iloc[colored_rows].reset_index(drop=True)

        if self.track_sheet_origin:
            filtered_df['__sheet_origin__'] = self.sheet_origin

        filtered_df['__color_D__'] = row_colors

        return filtered_df

    def get_dict_all_data(self) -> Dict[int, Dict[str, Any]]:
        """Возвращает весь словарь данных из self.data."""
        if self.data is None:
            print("Данные не загружены.")
            return {}

        return {index: row.to_dict() for index, row in self.data.iterrows()}

    def get_headers(self) -> List[str]:
        """Возвращает список заголовков колонок."""
        if self.data is None:
            return []
        return list(self.data.columns)

    def filter_and_save_columns(self, columns_to_save: Union[str, List[str], tuple]):
        """Сохраняет значения из указанных столбцов в словарь."""
        if self.data is None:
            print("Данные не загружены.")
            return

        if isinstance(columns_to_save, str):
            columns_to_save = [columns_to_save]
        elif not isinstance(columns_to_save, list):
            columns_to_save = list(columns_to_save)

        self.columns_save = columns_to_save
        self.filtered_data = {}

        for index, row in self.data.iterrows():
            self.filtered_data[index] = {
                col: row[col] for col in columns_to_save if col in row
            }

    def get_filtered_data(self) -> Optional[Dict[int, Dict[str, Any]]]:
        """Возвращает отфильтрованные данные."""
        return self.filtered_data

    @staticmethod
    def is_empty(val: Any) -> bool:
        """Проверяет, является ли значение пустым."""
        return (val is None or
                (isinstance(val, str) and val.strip().lower() in ("", "nan", "n/a", "none", "-", "null")) or
                (hasattr(pd, 'isna') and pd.isna(val)))

    def get_column_values(self, get_column: str,
                          foc_mode: bool = False,
                          skip_condition: Optional[callable] = None) -> List[str]:
        """Извлекает уникальные значения из указанной колонки."""
        if get_column not in self.columns_save:
            messagebox.showerror(
                "Ошибка",
                f"Название столбцов не совпадают.\n{get_column} в {self.columns_save}"
            )
            return []

        result = []
        all_data_dict = self.get_dict_all_data() if foc_mode else None

        for key in self.filtered_data:
            if foc_mode and skip_condition is not None:
                row_data = all_data_dict.get(key, {})
                if skip_condition(row_data):
                    continue

            value = self.filtered_data[key].get(get_column)
            if value is None or self.is_empty(value):
                continue

            value_str = str(value).strip()
            if not value_str:
                continue

            items = value_str.replace(", ", "|").split("|")

            for item in items:
                item = item.strip()
                if not item:
                    continue

                if ":" in item and "-" in item:
                    item = item[:len(item) // 2 + 1]

                if item not in result:
                    result.append(item)

        result.sort(reverse=True)
        result.append("")
        return result

    def return_data(self) -> Optional[pd.DataFrame]:
        """Возвращает загруженные данные."""
        return self.data


class MultiSheetReader:
    """Класс для чтения нескольких листов из одного Excel-файла."""

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.sheets: Dict[str, Optional[pd.DataFrame]] = {}

    def load_sheets(self, sheet_names: List[str]) -> Dict[str, Optional[pd.DataFrame]]:
        """Загружает указанные листы из Excel-файла."""
        for sheet_name in sheet_names:
            try:
                data = pd.read_excel(self.file_path, sheet_name=sheet_name)
                data = ExcelReader.fix_dataframe_encoding(data)
                self.sheets[sheet_name] = data
            except Exception as e:
                print(f"Ошибка при загрузке листа {sheet_name}: {e}")
                self.sheets[sheet_name] = None

        return self.sheets

    def get_sheet_as_dict(self, sheet_name: str) -> dict[Hashable, dict] | None:
        """Возвращает данные листа в виде словаря."""
        data = self.sheets.get(sheet_name)
        if data is None:
            return None

        return {index: row.to_dict() for index, row in data.iterrows()}

    def get_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """Возвращает DataFrame указанного листа."""
        return self.sheets.get(sheet_name)
