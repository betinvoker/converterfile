import sys, csv
import pandas as pd
import traceback
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from pathlib import Path
from datetime import datetime

class DesktopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Программа конвертации файлов")
        root.resizable(False, False)
        self.root.geometry("480x545")

        # Переменные для чекбоксов
        self.removing_spaces_var = tk.BooleanVar(value=False)  # По умолчанию выключен

        # Создаем основной фрейм
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка расширяемости окна
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(8, weight=1)

        # Кнопка
        self.butLoadFile = ttk.Button(main_frame, 
                                      text="Выбрать файлы для конвертации", 
                                      command=self.on_but_click_load_file
                                      )
        self.butLoadFile.grid(row=0, column=0, columnspan=2, pady=10, sticky="w")

        # Label
        self.format_label = ttk.Label(main_frame, text="Формат конвертации:")
        self.format_label.grid(row=1, column=0, padx=(0, 10), pady=5, sticky="w")

        # Выпадающий список (Combobox)
        self.format_combobox = ttk.Combobox(main_frame, 
                                           values=["CSV", "TXT", "XLSX", "UNL"],
                                           state="readonly")
        self.format_combobox.grid(row=1, column=1, pady=5, sticky="ew")
        self.format_combobox.set("TXT")  # Устанавливаем значение по умолчанию

        # Label
        self.encoding_label = ttk.Label(main_frame, text="Кодировка:")
        self.encoding_label.grid(row=2, column=0, padx=(0, 10), pady=5, sticky="w")

        # Выпадающий список (Combobox)
        self.encoding_combobox = ttk.Combobox(main_frame,
                                            values=["UTF-8", "ANSI", "cp1251"],
                                            state="readonly")
        self.encoding_combobox.grid(row=2, column=1, pady=5, sticky="ew")
        self.encoding_combobox.set("UTF-8")  # Устанавливаем значение по умолчанию

        # Первая строка: разделитель загружаемого файла
        self.delimiter_load_label = ttk.Label(main_frame, text="Разделитель загружаемого файла:")
        self.delimiter_load_label.grid(row=3, column=0, padx=(0, 10), pady=5, sticky="w")

        # Текстовое поле для разделителя загружаемого файла
        self.delimiter_load_entry = ttk.Entry(main_frame, width=3, justify='center')
        self.delimiter_load_entry.grid(row=3, column=1, pady=5, sticky="w")
        
        # Функция валидации для ограничения ввода одним символом
        def validate_input(char):
            return len(char) <= 1
        
        vcmd = (root.register(validate_input), '%P')
        self.delimiter_load_entry.configure(validate="key", validatecommand=vcmd)
        self.delimiter_load_entry.insert(0, "")  # Устанавливаем значение по умолчанию

       # Вторая строка: разделитель конвертируемого файла
        self.delimiter_convert_label = ttk.Label(main_frame, text="Разделитель конвертируемого файла:")
        self.delimiter_convert_label.grid(row=4, column=0, padx=(0, 10), pady=5, sticky="w")

        # Текстовое поле с ограничением в 1 символ
        self.delimiter_convert_entry = ttk.Entry(main_frame, width=3, justify='center')
        self.delimiter_convert_entry.grid(row=4, column=1, pady=5, sticky="w")
        
        vcmd = (root.register(validate_input), '%P')
        self.delimiter_convert_entry.configure(validate="key", validatecommand=vcmd)
        self.delimiter_convert_entry.insert(0, "")  # Устанавливаем значение по умолчанию

        # ЧЕКБОКСЫ - новая секция
        options_frame = ttk.LabelFrame(main_frame, text="Дополнительные опции", padding=5)
        options_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")
        
        # Чекбокс. Удаление лишних пробелов и скрытых символов
        self.removing_spaces_var_checkbox = ttk.Checkbutton(
            options_frame,
            text="Удаление лишних пробелов и скрытых символов",
            variable=self.removing_spaces_var,
            # command=self.on_options_change
        )
        self.removing_spaces_var_checkbox.grid(row=0, column=0, sticky="w", pady=2)

        # Кнопка
        self.butRun = ttk.Button(main_frame, 
                                 text="Выполнить конвертацию файлов", 
                                 command=self.on_btn_run_convert
                                 )
        self.butRun.grid(row=6, column=0, columnspan=2, pady=10, sticky="w")

         # Текстовая область для отображения файлов
        self.text_area = scrolledtext.ScrolledText( main_frame, 
                                                    wrap=tk.WORD, 
                                                    bg="#E3E3E3",
                                                    height=16,  # Уменьшаем высоту до 6 строк
                                                    width=50,  # Уменьшаем ширину
                                                    font=('Arial', 8))
        self.text_area.grid(row=7, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.text_area.insert(1.0, "Файлы не выбраны")
        self.text_area.config(state="disabled")

        # Инициализируем массив для хранения файлов
        self.array_of_paths = []

    def on_options_change(self):
        """Обработчик изменения состояния чекбоксов"""
        options = []
        if self.removing_spaces_var.get():
            options.append("* Удаление лишних пробелов и скрытых символов")

        # Выводим данные из array_of_paths в text_area (добавляем к существующему содержимому)
        self.text_area.config(state="normal")
        self.text_area.insert(tk.END, "\n")
        self.text_area.insert(tk.END, "-" * 50)
        self.text_area.insert(tk.END, f"\nВЫБРАННЫЕ ОПЦИИ:\n")
        self.text_area.insert(tk.END, "-" * 50)
        self.text_area.insert(tk.END, "\n" + "\n".join(options))
        self.text_area.config(state="disabled")

    def on_but_click_load_file(self):
        # Очищаем TEXT_AREA
        self.text_area.config(state="normal")
        self.text_area.delete(1.0, tk.END)
        self.text_area.config(state="disabled")

        # Выбор нескольких файлов
        file_paths = filedialog.askopenfilenames(
            title="Выберите файлы для конвертации",
            filetypes=[
                ("Все файлы", "*.*"),
                ("Текстовые файлы", "*.txt"),
                ("CSV файлы", "*.csv"),
                ("Excel файлы", "*.xlsx"),
                ("UNL файлы", "*.unl")
            ]
        )
                    
        if file_paths:
            # Включаем редактирование для записи
            self.text_area.config(state="normal")
            
            # Очищаем и заполняем текстовую область
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, f"Выбрано файлов: {len(file_paths)}\n\n")
            for i, file_path in enumerate(file_paths, 1):
                self.text_area.insert(tk.END, f"{i}. {file_path}\n")
            
            # Снова отключаем редактирование
            self.text_area.config(state="disabled")

    def on_btn_run_convert(self):
        try:
            # Очищаем массив перед заполнением
            self.array_of_paths = []

            # Получаем текст из текстовой области
            text_content = self.text_area.get(1.0, tk.END).strip()

            # Извлекаем пути из текстовой области
            lines = text_content.split('\n')
            for line in lines:
                if line.strip() and not line.startswith("Выбрано файлов:") and not line.startswith("Обработано файлов:"):
                    # Убираем нумерацию (1., 2., и т.д.)
                    file_path = line.split('. ', 1)[-1] if '. ' in line else line
                    if file_path and not file_path.startswith("Файлы не выбраны") \
                            and not file_path.startswith("Файл ") and not line.startswith("  Путь:") \
                            and not line.startswith("  Имя:") and not line.startswith("  Формат:") \
                            and not line.startswith("  Разделитель") and not line.startswith("  Целевой формат:") \
                            and not line.startswith("-" * 50) and not line.startswith("ВЫБРАННЫЕ ОПЦИИ:") \
                            and not line.startswith("*"):
                        # Создаем объект ConversionFile и добавляем в массив
                        conversion_file = self.ConversionFile(file_path)
                        self.array_of_paths.append(conversion_file)

            #Обработчик изменения состояния чекбоксов
            self.on_options_change()

            # Выводим данные из array_of_paths в text_area (добавляем к существующему содержимому)
            self.text_area.config(state="normal")
            # ЗАПИСЫВАЕМ ВРЕМЯ НАЧАЛА КОНВЕРТАЦИИ
            start_time = datetime.now()
            start_time_str = start_time.strftime("%Y-%m-%d %H:%M:%S")
            
            # Добавляем разделитель и новую информацию
            self.text_area.insert(tk.END, "\n" + "-" * 50 + "\n")
            self.text_area.insert(tk.END, "РЕЗУЛЬТАТЫ КОНВЕРТАЦИИ:\n")
            self.text_area.insert(tk.END, "-" * 50)

            if self.array_of_paths:            
                for i, conv_file in enumerate(self.array_of_paths, 1):
                    # Выполняем конвертацию и получаем результат
                    result_message = self.convert_file(i, conv_file)

                    self.text_area.insert(tk.END, f"{result_message}")
                    self.text_area.insert(tk.END, "-" * 50)
                
                self.text_area.insert(tk.END, f"\nОбработано файлов: {len(self.array_of_paths)}\n")
            else:
                self.text_area.insert(tk.END, "Нет файлов для обработки\n")


            # ЗАПИСЫВАЕМ ВРЕМЯ ЗАВЕРШЕНИЯ КОНВЕРТАЦИИ
            end_time = datetime.now()
            end_time_str = end_time.strftime("%Y-%m-%d %H:%M:%S")
            duration = end_time - start_time
            
            self.text_area.insert(tk.END, f"Время начала конвертации: {start_time_str}\n")
            self.text_area.insert(tk.END, f"Время завершения конвертации: {end_time_str}\n")
            self.text_area.insert(tk.END, f"Общее время выполнения: {duration.total_seconds():.2f} секунд\n")

        except Exception as e:
            self.text_area.config(state="normal")

            self.text_area.insert(tk.END, f"\n")
            self.text_area.insert(tk.END, "-" * 50)
            self.text_area.insert(tk.END, "\nПРОИЗОШЛА ОШИБКА:\n")
            self.text_area.insert(tk.END, "-" * 50)
            
            # Выводим основную информацию об ошибке
            self.text_area.insert(tk.END, f"\nТип ошибки: {type(e).__name__}")
            self.text_area.insert(tk.END, f"\nСообщение: {e}\n")

            # Получаем полную трассировку стека
            self.text_area.insert(tk.END, "-" * 50)
            self.text_area.insert(tk.END, "\nПОЛНАЯ ТРАССИРОВКА СТЕКА:\n")
            self.text_area.insert(tk.END, "-" * 50)
            self.text_area.insert(tk.END, f"\n{traceback.format_exc()}\n")

            # Дополнительная информация
            self.text_area.insert(tk.END, "-" * 50)
            self.text_area.insert(tk.END, "\nДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ:\n")
            self.text_area.insert(tk.END, "-" * 50)

            # Получаем информацию о последнем кадре стека
            tb = traceback.extract_tb(sys.exc_info()[2])
            if tb:
                last_frame = tb[-1]  # Последний кадр - где произошла ошибка
                self.text_area.insert(tk.END, f"\nФайл: {last_frame.filename}")
                self.text_area.insert(tk.END, f"\nСтрока: {last_frame.lineno}")
                self.text_area.insert(tk.END, f"\nФункция: {last_frame.name}")
                self.text_area.insert(tk.END, f"\nКод: {last_frame.line}\n")
                self.text_area.insert(tk.END, "-" * 50)

        finally:
            self.text_area.config(state="disabled")

    def convert_file(self, i, conv_file):
        try:
            # Получаем разделители
            delimiter_load = self.delimiter_load_entry.get() if self.delimiter_load_entry.get() else ','
            delimiter_convert = self.delimiter_convert_entry.get() if self.delimiter_convert_entry.get() else ','
            target_format = self.format_combobox.get().lower()
            encoding_format = self.encoding_combobox.get().lower()

            delimiter_display = self.delimiter_convert_entry.get() if target_format != "xlsx" else ""

            df = None

            if conv_file.format == 'xlsx':
                # Читаем Excel файл по полному пути 
                df = pd.read_excel(conv_file.path, dtype=str)
            elif conv_file.format in ['csv', 'txt', 'unl']:
                # Читаем CSV/TXT/UNL файл по полному пути
                # РУЧНОЕ ЧТЕНИЕ И АНАЛИЗ ФАЙЛА
                df = self._smart_file_reader(conv_file.path, delimiter_load, encoding_format)
            else:
                return f"\nФормат файла .{conv_file.format} не поддерживается\n"

            # ПРИМЕНЯЕМ УДАЛЕНИЕ ПРОБЕЛОВ КО ВСЕМ ФАЙЛАМ, если опция включена
            if self.removing_spaces_var.get() and df is not None:
                # 1. Обрабатываем заголовки столбцов
                df.columns = df.columns.str.strip()  # Удаляем пробелы по краям
                df.columns = df.columns.str.replace(r'\s+', ' ', regex=True)  # Заменяем множественные пробелы на один
                
                # 2. Обрабатываем данные в ячейках
                # Выбираем только строковые колонки
                string_columns = df.select_dtypes(include=['object']).columns

                for column in string_columns:
                    # Удаляем пробелы в начале и конце (включая первый символ если это пробел)
                    df[column] = df[column].str.strip()
                    # Заменяем множественные пробелы на один
                    df[column] = df[column].str.replace(r'\s+', ' ', regex=True)
         
            # Определяем расширение для выходного файла
            if target_format == 'xlsx':
                output_filename = f"{conv_file.name}_converted.xlsx"
                df.to_excel(output_filename, index=False)
            elif target_format in ['csv', 'txt', 'unl']:
                output_filename = f"{conv_file.name}_converted.{target_format}"
                df.to_csv(output_filename, sep=delimiter_convert, index=False,
                          encoding=encoding_format)
            else:
                return f"""
Файл {i}. Результат:
   Путь: {conv_file.path}
   Имя: {conv_file.name}
   Формат: {conv_file.format}
   Разделитель загрузки: '{self.delimiter_load_entry.get()}'
   Разделитель конвертации: '{delimiter_display}'
   Кодировка: '{self.encoding_combobox.get()}'
   Целевой формат: {self.format_combobox.get()}
   Неизвестный формат файла: {target_format}
"""

            return f"""
Файл {i}. Результат:
   Путь: {conv_file.path}
   Имя: {conv_file.name}
   Формат: {conv_file.format}
   Разделитель загрузки: '{self.delimiter_load_entry.get()}'
   Разделитель конвертации: '{delimiter_display}'
   Кодировка: '{self.encoding_combobox.get()}'
   Целевой формат: {self.format_combobox.get()}
   Результат: Успешно сконвертирован в {output_filename}
"""
        except Exception as e:
            return f"""
Файл {i}. Результат:
   Путь: {conv_file.path}
   Имя: {conv_file.name}
   Формат: {conv_file.format}
   Разделитель загрузки: '{self.delimiter_load_entry.get()}'
   Разделитель конвертации: '{delimiter_display}'
   Кодировка: '{self.encoding_combobox.get()}'
   Целевой формат: {self.format_combobox.get()}
   Ошибка конвертации: {str(e)}
"""

    def _smart_file_reader(self, file_path, delimiter, encoding):
        """Умное чтение файла с автоматическим определением структуры"""
        with open(file_path, 'r', encoding=encoding) as f:
            lines = f.readlines()
        
        # Функция для проверки является ли строка разделителем
        def is_delimiter_line(line):
            if not line.strip():
                return False
            # Проверяем, содержит ли строка только символы разделителей
            return all(char in '-=| ' for char in line.strip())
        
        # Функция для разбора строки на колонки
        def parse_line(line):
            return [part.strip() for part in line.split(delimiter)]
        
        # Ищем структуру файла
        header_found = False
        headers = None
        data_rows = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Пропускаем строки-разделители
            if is_delimiter_line(line):
                continue
                
            parts = parse_line(line)
            
            # Если еще не нашли заголовки
            if not header_found:
                # Проверяем, похожа ли строка на заголовки
                # Заголовки обычно содержат текст (буквы), а не только цифры
                looks_like_header = any(
                    any(c.isalpha() for c in part) for part in parts
                ) if parts else False
                
                if looks_like_header:
                    headers = parts
                    header_found = True
                else:
                    # Если первая не-разделительная строка не похожа на заголовки, 
                    # но содержит данные - создаем автоматические заголовки
                    if len(parts) > 1:
                        headers = [f"Column_{i+1}" for i in range(len(parts))]
                        header_found = True
                        # Добавляем текущую строку как данные
                        data_rows.append(parts)
            else:
                # После нахождения заголовков добавляем данные
                if len(parts) == len(headers):
                    data_rows.append(parts)
                elif len(parts) > len(headers):
                    # Если больше колонок, обрезаем до размера заголовков
                    data_rows.append(parts[:len(headers)])
                else:
                    # Если меньше колонок, дополняем пустыми значениями
                    padded_parts = parts + [''] * (len(headers) - len(parts))
                    data_rows.append(padded_parts)
        
        # Если не нашли заголовки, но есть данные
        if not header_found and data_rows:
            headers = [f"Column_{i+1}" for i in range(len(data_rows[0]))]
        
        # Создаем DataFrame
        if headers and data_rows:
            return pd.DataFrame(data_rows, columns=headers)
        elif headers:
            return pd.DataFrame(columns=headers)
        else:
            # Если ничего не нашли, возвращаем пустой DataFrame
            return pd.DataFrame()

    class ConversionFile:
         def __init__(self, path):
            self.path = path
            self.name = Path(path).stem
            if '.' in path:
                self.format = path.rsplit('.', 1)[1].lower()
            else:
                self.format = "Нет расширения"

def main():
    root = tk.Tk()
    app = DesktopApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()