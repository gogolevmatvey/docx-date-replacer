"""
Модуль для поиска и замены дат в текстах.
"""

import re
from typing import Optional, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from docx.document import Document as DocumentType
else:
    DocumentType = None

from .config import DATE_PATTERN


class DateReplacer:
    """Класс для поиска и замены даты «29» января 2026 г. на новую дату."""
    
    # Регулярное выражение для поиска ТОЛЬКО даты «29» января 2026 г.
    # С пробелами и без, с "г." и "г"
    # Используется паттерн из config.py
    _DATE_PATTERN = re.compile(DATE_PATTERN)
    
    # Варианты написания даты для поиска
    OLD_DATE_VARIANTS = [
        "«29» января 2026 г.",
        "«29» января 2026г.",
        "«29»  января  2026 г.",
        "«29»  января  2026г.",
        "« 29 » января 2026 г.",
        "« 29 » января 2026г.",
        # Вариант с двойным пробелом перед годом (как в ФОМ_090303_Алгорит_прогр_Шутов)
        "«29» января  2026 г.",
        "«29» января  2026г.",
    ]
    
    # Названия месяцев для валидации
    VALID_MONTHS = {
        'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
        'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
    }
    
    def __init__(self, new_date: str):
        """
        Инициализация заменщика дат.

        Args:
            new_date: Новая дата для замены (например, "«26» февраля 2026 г.")
        """
        self.new_date = new_date
        self.old_date_variants = self.OLD_DATE_VARIANTS

    def find_date(self, text: str) -> Optional[str]:
        """
        Поиск даты «29» января 2026 г. в тексте.

        Args:
            text: Текст для поиска

        Returns:
            Найденная дата или None, если не найдено
        """
        match = self._DATE_PATTERN.search(text)
        if match:
            return match.group(0)
        return None

    def find_date_in_first_paragraphs(self, doc: "DocumentType", max_paragraphs: int = 50) -> bool:
        """
        Проверка наличия даты «29» января 2026 г. на первой странице.

        Args:
            doc: Объект Document
            max_paragraphs: Максимальное количество параграфов для проверки

        Returns:
            True если дата найдена на первой странице
        """
        from docx.oxml.ns import qn

        # Проверяем первые N параграфов
        for i, paragraph in enumerate(doc.paragraphs[:max_paragraphs]):
            if paragraph.text.strip():
                if self.find_date(paragraph.text):
                    return True

        # Проверяем первые ячейки таблиц (для документов с таблицей в начале)
        cell_count = 0
        for tc in doc._element.body.iter(qn('w:tc')):
            cell_text = ''
            for t in tc.iter(qn('w:t')):
                if t.text:
                    cell_text += t.text
            if cell_text.strip():
                if self.find_date(cell_text):
                    return True
                cell_count += 1
                if cell_count > 10:  # Проверяем только первые несколько ячеек
                    break

        return False

    def find_date_details(self, text: str) -> Optional[Tuple[str, str, str]]:
        """
        Поиск даты с возвратом деталей (день, месяц, год).

        Args:
            text: Текст для поиска

        Returns:
            Кортеж (день, месяц, год) или None
        """
        match = self._DATE_PATTERN.search(text)
        if match:
            # match.group(0) - вся дата, group(1) - месяц, group(2) - год
            return ("29", match.group(1), match.group(2))
        return None
    
    def replace_date(self, text: str) -> Tuple[str, bool]:
        """
        Замена даты в тексте на новую.

        Args:
            text: Текст для обработки

        Returns:
            Кортеж (обработанный текст, флаг изменения)
        """
        new_text, count = self._DATE_PATTERN.subn(self.new_date, text)
        return (new_text, count > 0)

    def validate_date_format(self, date_string: str) -> bool:
        """
        Валидация формата даты.

        Args:
            date_string: Строка для проверки

        Returns:
            True если формат корректен
        """
        match = self._DATE_PATTERN.match(date_string.strip())
        if not match:
            return False

        # match.groups() возвращает (месяц, год) - день захардкожен в паттерне
        month, year = match.groups()

        # Проверка месяца
        if month.lower() not in self.VALID_MONTHS:
            return False

        # Проверка года (4 цифры)
        if len(year) != 4:
            return False

        return True
    
    def has_approval_block(self, text: str) -> bool:
        """
        Проверка наличия блока «УТВЕРЖДАЮ» в тексте.
        
        Args:
            text: Текст для проверки
            
        Returns:
            True если блок найден
        """
        return "УТВЕРЖДАЮ" in text.upper()
