"""
Тесты для модуля date_replacer.
"""

import unittest
import sys
import os

# Добавляем корневую папку в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.date_replacer import DateReplacer


class TestDateReplacer(unittest.TestCase):
    """Тесты для класса DateReplacer."""

    def setUp(self):
        """Настройка перед каждым тестом."""
        self.replacer = DateReplacer("«26» февраля 2026 г.")

    def test_find_date_standard_format(self):
        """Поиск даты в стандартном формате."""
        text = "«29» января 2026 г."
        result = self.replacer.find_date(text)
        self.assertEqual(result, "«29» января 2026 г.")

    def test_find_date_without_space_before_g(self):
        """Поиск даты без пробела перед 'г'."""
        text = "«29» января 2026г."
        result = self.replacer.find_date(text)
        self.assertEqual(result, "«29» января 2026г.")

    def test_find_date_with_multiple_spaces(self):
        """Поиск даты с множественными пробелами."""
        text = "«29»  января  2026 г."
        result = self.replacer.find_date(text)
        self.assertEqual(result, "«29»  января  2026 г.")

    def test_find_date_not_found(self):
        """Поиск даты, когда её нет в тексте."""
        text = "«27» февраля 2026 г."
        result = self.replacer.find_date(text)
        self.assertIsNone(result)

    def test_find_date_wrong_day(self):
        """Поиск даты с неправильным днём."""
        text = "«28» января 2026 г."
        result = self.replacer.find_date(text)
        self.assertIsNone(result)

    def test_find_date_wrong_month(self):
        """Поиск даты с неправильным месяцем."""
        text = "«29» февраля 2026 г."
        result = self.replacer.find_date(text)
        self.assertIsNone(result)

    def test_find_date_wrong_year(self):
        """Поиск даты с неправильным годом."""
        text = "«29» января 2025 г."
        result = self.replacer.find_date(text)
        self.assertIsNone(result)

    def test_find_date_details(self):
        """Поиск даты с возвратом деталей."""
        text = "«29» января 2026 г."
        result = self.replacer.find_date_details(text)
        self.assertEqual(result, ("29", "января", "2026"))

    def test_find_date_details_not_found(self):
        """Поиск даты с возвратом деталей, когда даты нет."""
        text = "«27» февраля 2026 г."
        result = self.replacer.find_date_details(text)
        self.assertIsNone(result)

    def test_replace_date_standard(self):
        """Замена даты в стандартном формате."""
        text = "Документ от «29» января 2026 г."
        new_text, changed = self.replacer.replace_date(text)
        self.assertTrue(changed)
        self.assertEqual(new_text, "Документ от «26» февраля 2026 г.")

    def test_replace_date_no_change(self):
        """Замена даты, когда её нет в тексте."""
        text = "Документ от «27» февраля 2026 г."
        new_text, changed = self.replacer.replace_date(text)
        self.assertFalse(changed)
        self.assertEqual(new_text, text)

    def test_replace_date_multiple_occurrences(self):
        """Замена нескольких дат в тексте."""
        text = "«29» января 2026 г. и ещё «29» января 2026 г."
        new_text, changed = self.replacer.replace_date(text)
        self.assertTrue(changed)
        self.assertEqual(new_text, "«26» февраля 2026 г. и ещё «26» февраля 2026 г.")

    def test_validate_date_format_valid(self):
        """Валидация корректной даты."""
        text = "«29» января 2026 г."
        result = self.replacer.validate_date_format(text)
        self.assertTrue(result)

    def test_validate_date_format_invalid_day(self):
        """Валидация даты с неправильным днём."""
        text = "«32» января 2026 г."
        result = self.replacer.validate_date_format(text)
        # Паттерн найдёт дату, но день 32 невалиден
        # Однако наш паттерн ищет только «29», так что это не найдётся
        self.assertFalse(result)

    def test_validate_date_format_invalid_month(self):
        """Валидация даты с неправильным месяцем."""
        text = "«29» мартобря 2026 г."
        result = self.replacer.validate_date_format(text)
        self.assertFalse(result)

    def test_has_approval_block(self):
        """Проверка наличия блока «УТВЕРЖДАЮ»."""
        text = "УТВЕРЖДАЮ: Директор"
        result = self.replacer.has_approval_block(text)
        self.assertTrue(result)

    def test_has_approval_block_lowercase(self):
        """Проверка наличия блока «УТВЕРЖДАЮ» в нижнем регистре."""
        text = "утверждаю: директор"
        result = self.replacer.has_approval_block(text)
        self.assertTrue(result)

    def test_has_approval_block_not_found(self):
        """Проверка отсутствия блока «УТВЕРЖДАЮ»."""
        text = "СОГЛАСОВАНО: Директор"
        result = self.replacer.has_approval_block(text)
        self.assertFalse(result)


if __name__ == "__main__":
    unittest.main()
