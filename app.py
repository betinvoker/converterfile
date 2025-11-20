import sys
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
        self.root.geometry("480x600")

        # Создаем основной фрейм
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка расширяемости окна
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

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
                                           values=["CSV", "TXT", "XLSX"],
                                           state="readonly")
        self.format_combobox.grid(row=1, column=1, pady=5, sticky="ew")
        self.format_combobox.set("TXT")  # Устанавливаем значение по умолчанию

        # Первая строка: разделитель загружаемого файла
        self.delimiter_load_label = ttk.Label(main_frame, text="Разделитель загружаемого файла:")
        self.delimiter_load_label.grid(row=2, column=0, padx=(0, 10), pady=5, sticky="w")

        # Текстовое поле для разделителя загружаемого файла
        self.delimiter_load_entry = ttk.Entry(main_frame, width=3, justify='center')
        self.delimiter_load_entry.grid(row=2, column=1, pady=5, sticky="w")
        
        # Функция валидации для ограничения ввода одним символом
        def validate_input(char):
            return len(char) <= 1
        
        vcmd = (root.register(validate_input), '%P')
        self.delimiter_load_entry.configure(validate="key", validatecommand=vcmd)
        self.delimiter_load_entry.insert(0, "")  # Устанавливаем значение по умолчанию

       # Вторая строка: разделитель конвертируемого файла
        self.delimiter_convert_label = ttk.Label(main_frame, text="Разделитель конвертируемого файла:")
        self.delimiter_convert_label.grid(row=3, column=0, padx=(0, 10), pady=5, sticky="w")

        # Текстовое поле с ограничением в 1 символ
        self.delimiter_convert_entry = ttk.Entry(main_frame, width=3, justify='center')
        self.delimiter_convert_entry.grid(row=3, column=1, pady=5, sticky="w")
        
        vcmd = (root.register(validate_input), '%P')
        self.delimiter_convert_entry.configure(validate="key", validatecommand=vcmd)
        self.delimiter_convert_entry.insert(0, "")  # Устанавливаем значение по умолчанию

        # Кнопка
        self.butRun = ttk.Button(main_frame, 
                                 text="Выполнить конвертацию файлов", 
                                 command=self.on_btn_run_convert
                                 )
        self.butRun.grid(row=4, column=0, columnspan=2, pady=10, sticky="w")

         # Текстовая область для отображения файлов
        self.text_area = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, bg="#E3E3E3")
        self.text_area.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.text_area.insert(1.0, "Файлы не выбраны")
        self.text_area.config(state="disabled")

        # Инициализируем массив для хранения файлов
        self.array_of_paths = []

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
                ("Excel файлы", "*.xlsx")
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
                            and not line.startswith("-" * 40):
                        # Создаем объект ConversionFile и добавляем в массив
                        conversion_file = self.ConversionFile(file_path)
                        self.array_of_paths.append(conversion_file)

            # Выводим данные из array_of_paths в text_area (добавляем к существующему содержимому)
            self.text_area.config(state="normal")
            # ЗАПИСЫВАЕМ ВРЕМЯ НАЧАЛА КОНВЕРТАЦИИ
            start_time = datetime.now()
            start_time_str = start_time.strftime("%Y-%m-%d %H:%M:%S")
            
            # Добавляем разделитель и новую информацию
            self.text_area.insert(tk.END, "\n" + "-" * 50 + "\n")
            self.text_area.insert(tk.END, "РЕЗУЛЬТАТЫ КОНВЕРТАЦИИ:\n")
            self.text_area.insert(tk.END, "-" * 50 + "\n")

            if self.array_of_paths:            
                for i, conv_file in enumerate(self.array_of_paths, 1):
                    # Выполняем конвертацию и получаем результат
                    result_message = self.convert_file(conv_file)

                    self.text_area.insert(tk.END, f"Файл {i}:\n")
                    self.text_area.insert(tk.END, f"  Путь: {conv_file.path}\n")
                    self.text_area.insert(tk.END, f"  Имя: {conv_file.name}\n")
                    self.text_area.insert(tk.END, f"  Формат: {conv_file.format}\n")
                    self.text_area.insert(tk.END, f"  Разделитель загрузки: '{self.delimiter_load_entry.get()}'\n")
                    self.text_area.insert(tk.END, f"  Разделитель конвертации: '{self.delimiter_convert_entry.get()}'\n")
                    self.text_area.insert(tk.END, f"  Целевой формат: {self.format_combobox.get()}\n")
                    self.text_area.insert(tk.END, f"  Результат: {result_message}\n")
                    self.text_area.insert(tk.END, "-" * 50 + "\n")
                
                self.text_area.insert(tk.END, f"Обработано файлов: {len(self.array_of_paths)}\n")
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

    def convert_file(self, conv_file):
        try:
            # Получаем разделители
            delimiter_load = self.delimiter_load_entry.get() if self.delimiter_load_entry.get() else ','
            delimiter_convert = self.delimiter_convert_entry.get() if self.delimiter_convert_entry.get() else ','
            target_format = self.format_combobox.get().lower()

            if conv_file.format == 'xlsx':
                # Читаем Excel файл по полному пути
                df = pd.read_excel(conv_file.path, dtype=str)
                output_filename = f"{conv_file.name}_converted.xlsx"
                df.to_excel(output_filename, index=False)
                return f"Успешно конвертирован в {output_filename}"
            elif conv_file.format in ['csv', 'txt']:
                # Читаем CSV/TXT файл по полному пути
                df = pd.read_csv(conv_file.path, sep=delimiter_load, dtype=str)
            else:
                return f"Формат файла .{conv_file.format} не поддерживается"
                
            # Определяем расширение для выходного файла
            if target_format == 'xlsx':
                output_filename = f"{conv_file.name}_converted.xlsx"
                df.to_excel(output_filename, index=False)
            elif target_format in ['csv', 'txt']:
                output_filename = f"{conv_file.name}_converted.{target_format}"
                df.to_csv(output_filename, sep=delimiter_convert, index=False)   
            else:
                return f"Неизвестный формат файла: {target_format}"

            return f"Успешно конвертирован в {output_filename}"
        except Exception as e:
            return f"Ошибка конвертации: {str(e)}"

    class ConversionFile:
         def __init__(self, path):
            self.path = path    # имя человека
            self.name = Path(path).stem
            # Безопасное получение расширения
            if '.' in path:
                self.format = path.rsplit('.', 1)[1].lower()
            else:
                self.format = "нет расширения"

def main():
    root = tk.Tk()
    app = DesktopApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()