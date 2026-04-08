"""
Тесты для модуля date_replacer.
"""

import unittest

from src.date_replacer import DateReplacer


class TestDateReplacer(unittest.TestCase):
    """Тесты для класса DateReplacer."""

    def setUp(self):
        """Настройка перед каждым тестом."""
        self.replacer = DateReplacer("«29» января 2026 г.", "«26» февраля 2026 г.")

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

    def test_search_pattern_returns_compiled_regex(self):
        """Проверка что search_pattern возвращает скомпилированный regex."""
        import re
        pattern = self.replacer.search_pattern()
        self.assertIsInstance(pattern, re.Pattern)
        self.assertTrue(pattern.search("«29» января 2026 г."))

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


if __name__ == "__main__":
    unittest.main()
