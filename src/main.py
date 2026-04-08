"""
Точка входа приложения Docx Date Replacer.
"""

import sys


def main():
    """Запуск приложения."""
    try:
        from src.gui import App
        app = App()
        app.mainloop()
    except ImportError as e:
        print(f"Не удалось запустить GUI: {e}")
        print("Установите зависимости: pip install -e .")
        sys.exit(1)


if __name__ == "__main__":
    main()
