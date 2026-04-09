"""
Модуль графического интерфейса (CustomTkinter).
"""

import os
import logging
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Импортируем логику
from .date_replacer import DateReplacer
from .docx_processor import DocxProcessor
from .config import DEFAULT_CONFIG

logger = logging.getLogger(__name__)


class App(ctk.CTk):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()

        # Настройки окна
        self.title("Замена даты в документах")
        self.geometry("700x650")
        self.minsize(600, 500)

        # Иконка окна
        import sys
        if getattr(sys, 'frozen', False):
            # Запуск из .exe — иконка во временной папке _MEIPASS
            base_dir = sys._MEIPASS  # type: ignore[attr-defined]
        else:
            # Запуск из исходников — корень проекта
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        icon_path = os.path.join(base_dir, "docx-date-replacer.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.is_processing = False

        # UI
        self.create_widgets()

    def create_widgets(self):
        """Создание элементов интерфейса."""
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(12, weight=1)

        # --- Заголовок ---
        self.title_label = ctk.CTkLabel(self, text="Замена даты в документах", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.grid(row=0, column=0, columnspan=3, padx=20, pady=(20, 10), sticky="w")

        # --- Исходная папка ---
        self.source_label = ctk.CTkLabel(self, text="Исходная папка:", anchor="w")
        self.source_label.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")

        self.source_entry = ctk.CTkEntry(self, placeholder_text="Выберите папку...")
        self.source_entry.grid(row=2, column=0, columnspan=2, padx=20, pady=5, sticky="ew")

        self.source_btn = ctk.CTkButton(self, text="Обзор...", width=80, command=self.browse_source)
        self.source_btn.grid(row=2, column=2, padx=20, pady=5, sticky="e")

        # --- Папка вывода ---
        self.output_label = ctk.CTkLabel(self, text="Папка вывода:", anchor="w")
        self.output_label.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")

        self.output_entry = ctk.CTkEntry(self, placeholder_text="Выберите папку...")
        self.output_entry.grid(row=4, column=0, columnspan=2, padx=20, pady=5, sticky="ew")

        self.output_btn = ctk.CTkButton(self, text="Обзор...", width=80, command=self.browse_output)
        self.output_btn.grid(row=4, column=2, padx=20, pady=5, sticky="e")

        # --- Даты ---
        self.old_date_label = ctk.CTkLabel(self, text="Дата для замены (что ищем):", anchor="w")
        self.old_date_label.grid(row=5, column=0, padx=20, pady=(15, 0), sticky="w")

        self.old_date_entry = ctk.CTkEntry(self)
        self.old_date_entry.grid(row=6, column=0, columnspan=3, padx=20, pady=5, sticky="ew")
        self.old_date_entry.insert(0, DEFAULT_CONFIG.old_date)

        self.new_date_label = ctk.CTkLabel(self, text="Новая дата (на что меняем):", anchor="w")
        self.new_date_label.grid(row=7, column=0, padx=20, pady=(10, 0), sticky="w")

        self.new_date_entry = ctk.CTkEntry(self)
        self.new_date_entry.grid(row=8, column=0, columnspan=3, padx=20, pady=5, sticky="ew")
        self.new_date_entry.insert(0, DEFAULT_CONFIG.new_date)

        # --- Кнопка запуска ---
        self.start_btn = ctk.CTkButton(self, text="▶ Запустить обработку", height=40, font=ctk.CTkFont(size=16), command=self.start_processing)
        self.start_btn.grid(row=9, column=0, columnspan=3, padx=20, pady=20, sticky="ew")

        # --- Прогресс бар ---
        self.progress = ctk.CTkProgressBar(self)
        self.progress.grid(row=10, column=0, columnspan=3, padx=20, pady=5, sticky="ew")
        self.progress.set(0)

        # --- Статус ---
        self.status_label = ctk.CTkLabel(self, text="Готово к работе", anchor="w")
        self.status_label.grid(row=11, column=0, columnspan=3, padx=20, pady=5, sticky="w")

        # --- Лог (Textbox) ---
        self.log_box = ctk.CTkTextbox(self, height=150, state="normal")
        self.log_box.grid(row=12, column=0, columnspan=3, padx=20, pady=(20, 5), sticky="nsew")

        # Кнопка для копирования логов
        self.copy_btn = ctk.CTkButton(self, text="📋 Копировать логи", command=self.copy_logs)
        self.copy_btn.grid(row=13, column=0, columnspan=3, padx=20, pady=(5, 20), sticky="ew")

    def copy_logs(self):
        """Копирование содержимого лога в буфер обмена."""
        logs = self.log_box.get("1.0", "end")
        self.clipboard_clear()
        self.clipboard_append(logs)
        self.status_label.configure(text="Логи скопированы в буфер обмена!")

    def browse_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_entry.delete(0, 'end')
            self.source_entry.insert(0, folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, 'end')
            self.output_entry.insert(0, folder)

    def log(self, message):
        """Добавление сообщения в лог."""
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.update()

    def start_processing(self):
        """Запуск обработки в отдельном потоке."""
        if self.is_processing:
            return

        # Валидация
        source = self.source_entry.get()
        output = self.output_entry.get()
        old_date = self.old_date_entry.get()
        new_date = self.new_date_entry.get()

        if not os.path.exists(source):
            messagebox.showerror("Ошибка", f"Исходная папка не найдена:\n{source}")
            return

        if not output:
            messagebox.showerror("Ошибка", "Не указана папка вывода!")
            return

        if not old_date:
            messagebox.showerror("Ошибка", "Не указана дата для замены!")
            return

        if not new_date:
            messagebox.showerror("Ошибка", "Не указана новая дата!")
            return

        self.is_processing = True
        self.start_btn.configure(state="disabled")
        self.progress.set(0)
        self.log_box.delete("1.0", "end")

        # Запуск в потоке
        thread = threading.Thread(target=self.run_task, args=(source, output, old_date, new_date), daemon=True)
        thread.start()

    def run_task(self, source, output, old_date, new_date):
        """Основная задача (выполняется в фоне)."""
        try:
            self.log("🚀 Запуск обработки...")
            self.log(f"🔍 Ищем: {old_date}")
            self.log(f"📝 Меняем на: {new_date}")
            self.status_label.configure(text="Обработка...")

            # Настройка логирования для вывода в GUI
            class GUILogHandler(logging.Handler):
                def __init__(self, log_func):
                    super().__init__()
                    self.log_func = log_func

                def emit(self, record):
                    msg = self.format(record)
                    self.log_func(msg)

            gui_handler = GUILogHandler(self.log)
            gui_handler.setLevel(logging.DEBUG)

            for name in ['src.date_replacer', 'src.docx_processor']:
                log = logging.getLogger(name)
                log.addHandler(gui_handler)

            # Настройка процессора
            replacer = DateReplacer(old_date, new_date)
            processor = DocxProcessor(replacer)

            # Поиск файлов
            files = processor.find_docx_files(source)
            total = len(files)
            self.log(f"📂 Найдено файлов: {total}")

            if total == 0:
                self.log("⚠️ Файлы .docx не найдены.")
                self.status_label.configure(text="Файлы не найдены")
                self.on_finish()
                return

            # Создание папки вывода
            processor.copy_folder_structure(source, output)
            self.log(f"📂 Структура папок создана в: {output}")

            # Обработка
            replaced_total = 0

            for i, filepath in enumerate(files, 1):
                filename = os.path.basename(filepath)
                out_path = processor.get_output_path(filepath, source, output)

                self.log(f"[{i}/{total}] {filename}...")

                success, message, count = processor.process_document(filepath, out_path)

                if success:
                    replaced_total += count
                    if count > 0:
                        self.log(f"  ✅ {message}")
                    else:
                        self.log(f"  ℹ️ Скопировано без изменений")
                else:
                    self.log(f"  ❌ Ошибка: {message}")

                progress = i / total
                self.progress.set(progress)

            self.log("-" * 40)
            self.log(f"🎉 Готово! Всего замен: {replaced_total}")
            self.status_label.configure(text="Завершено успешно")

        except Exception as e:
            self.log(f"❌ Критическая ошибка: {str(e)}")
            self.status_label.configure(text="Ошибка!")
        finally:
            self.on_finish()

    def on_finish(self):
        """Завершение задачи."""
        self.is_processing = False
        self.start_btn.configure(state="normal")
        self.progress.set(1.0)
