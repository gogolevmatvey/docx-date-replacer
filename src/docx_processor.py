"""
Модуль для обработки документов .docx.
"""

import os
from typing import List, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from docx.document import Document as DocumentType
else:
    DocumentType = None

from docx import Document

from .date_replacer import DateReplacer


class DocxProcessor:
    """Класс для обработки документов .docx."""
    
    def __init__(self, date_replacer: DateReplacer):
        """
        Инициализация процессора документов.
        
        Args:
            date_replacer: Экземпляр DateReplacer для замены дат
        """
        self.date_replacer = date_replacer
    
    def load_document(self, filepath: str) -> "DocumentType":
        """
        Загрузка документа .docx.
        
        Args:
            filepath: Путь к файлу
            
        Returns:
            Объект Document
            
        Raises:
            FileNotFoundError: Если файл не найден
            Exception: Если файл повреждён
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Файл не найден: {filepath}")
        return Document(filepath)
    
    def save_document(self, doc: "DocumentType", filepath: str) -> None:
        """
        Сохранение документа .docx.
        
        Args:
            doc: Объект Document
            filepath: Путь для сохранения
        """
        # Создаём директорию, если не существует
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        doc.save(filepath)
    
    def process_paragraphs(self, doc: "DocumentType", first_page_only: bool = True) -> Tuple[int, int]:
        """
        Обработка параграфов документа с сохранением форматирования.

        Args:
            doc: Объект Document
            first_page_only: Обрабатывать только первую страницу

        Returns:
            Кортеж (количество обработанных параграфов, количество замен)
        """
        processed = 0
        replaced = 0

        # Определяем диапазон параграфов (первые 50 для первой страницы)
        max_paragraphs = 50 if first_page_only else len(doc.paragraphs)
        paragraphs_to_process = doc.paragraphs[:max_paragraphs]

        for paragraph in paragraphs_to_process:
            if paragraph.text.strip():
                # Проверяем, есть ли в параграфе дата «29» января 2026 г.
                if self.date_replacer.find_date(paragraph.text):
                    # Собираем все run и их позиции
                    runs_with_text = []
                    for run in paragraph.runs:
                        if run.text is not None:
                            runs_with_text.append(run)

                    # Ищем дату в объединённом тексте
                    full_text = paragraph.text
                    
                    for old_var in self.date_replacer.old_date_variants:
                        if old_var in full_text:
                            # Находим позицию даты в полном тексте
                            start_idx = full_text.find(old_var)
                            end_idx = start_idx + len(old_var)
                            new_date = self.date_replacer.new_date

                            # Определяем, какие run покрывают дату
                            current_pos = 0
                            run_ranges = []
                            for run in runs_with_text:
                                run_start = current_pos
                                run_end = current_pos + len(run.text)
                                if run_start < end_idx and run_end > start_idx:
                                    # Этот run пересекается с датой
                                    overlap_start = max(run_start, start_idx)
                                    overlap_end = min(run_end, end_idx)
                                    run_ranges.append((run, overlap_start - run_start, overlap_end - run_start))
                                current_pos = run_end

                            # Заменяем части даты в соответствующих run
                            new_date_pos = 0
                            for run, rel_start, rel_end in run_ranges:
                                if run.text is not None:
                                    # Вычисляем, какую часть новой даты вставить
                                    new_part = new_date[new_date_pos:new_date_pos + (rel_end - rel_start)]
                                    old_text = run.text
                                    run.text = old_text[:rel_start] + new_part + old_text[rel_end:]
                                    new_date_pos += (rel_end - rel_start)

                            replaced += 1
                            break

                    processed += 1
                    if first_page_only and replaced > 0:
                        # Нашли и заменили дату на первой странице
                        break

        return (processed, replaced)

    def process_tables(self, doc: "DocumentType", first_page_only: bool = True) -> Tuple[int, int]:
        """
        Обработка таблиц документа (только первые ячейки для первой страницы).

        Args:
            doc: Объект Document
            first_page_only: Обрабатывать только первую страницу

        Returns:
            Кортеж (количество обработанных ячеек, количество замен)
        """
        from docx.oxml.ns import qn
        from copy import deepcopy

        processed = 0
        replaced = 0
        max_cells = 20 if first_page_only else 1000  # Ограничиваем количество ячеек

        # Проходим по первым ячейкам таблиц (первая страница)
        for tc in doc._element.body.iter(qn('w:tc')):
            if processed >= max_cells:
                break

            # Собираем весь текст из ячейки
            text_parts = []
            text_elements = []
            for t in tc.iter(qn('w:t')):
                if t.text is not None:
                    text_parts.append(t.text)
                    text_elements.append(t)

            cell_text = ''.join(text_parts)

            if cell_text.strip():
                # Проверяем наличие ТОЛЬКО даты «29» января 2026 г.
                for old_var in self.date_replacer.old_date_variants:
                    if old_var in cell_text:
                        start_idx = cell_text.find(old_var)
                        end_idx = start_idx + len(old_var)
                        new_date = self.date_replacer.new_date

                        # Находим элементы, которые покрывают дату
                        current_pos = 0
                        date_elements = []
                        for t in text_elements:
                            elem_start = current_pos
                            elem_end = current_pos + len(t.text)
                            if elem_start < end_idx and elem_end > start_idx:
                                date_elements.append((t, elem_start, elem_end))
                            current_pos = elem_end

                        # Заменяем текст в элементах с сохранением форматирования
                        if date_elements:
                            # Копируем rPr из первого элемента даты (если есть) для сохранения шрифта
                            first_elem = date_elements[0][0]
                            first_r = first_elem.getparent()
                            orig_rPr = None
                            orig_rFonts = None
                            
                            if first_r is not None:
                                rPr = first_r.find(qn('w:rPr'))
                                if rPr is not None:
                                    orig_rPr = deepcopy(rPr)
                                    orig_rFonts = rPr.find(qn('w:rFonts'))
                                    if orig_rFonts is not None:
                                        orig_rFonts = deepcopy(orig_rFonts)
                            
                            # Создаём НОВЫЙ w:r с правильной датой и форматированием
                            from docx.oxml import OxmlElement
                            
                            new_r = OxmlElement('w:r')
                            
                            # Создаём w:rPr с явным указанием размера 12pt
                            rPr_new = OxmlElement('w:rPr')
                            
                            # Копируем шрифт из оригинала (если есть)
                            if orig_rFonts is not None:
                                rPr_new.append(deepcopy(orig_rFonts))
                            else:
                                # Устанавливаем Times New Roman по умолчанию
                                rFonts = OxmlElement('w:rFonts')
                                rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', 'Times New Roman')
                                rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', 'Times New Roman')
                                rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs', 'Times New Roman')
                                rPr_new.append(rFonts)
                            
                            # Устанавливаем размер шрифта 24 (12pt) - ЯВНО с высоким приоритетом
                            sz = OxmlElement('w:sz')
                            sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '24')
                            rPr_new.append(sz)
                            
                            szCs = OxmlElement('w:szCs')
                            szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '24')
                            rPr_new.append(szCs)
                            
                            new_r.append(rPr_new)
                            
                            # Создаём w:t с новой датой
                            new_t = OxmlElement('w:t')
                            new_t.text = new_date
                            new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            new_r.append(new_t)
                            
                            # Вставляем новый w:r ПЕРЕД первым элементом даты
                            first_r.addprevious(new_r)
                            
                            # Удаляем ВСЕ элементы, которые были частью даты
                            for t, elem_start, elem_end in date_elements:
                                r = t.getparent()
                                # Очищаем текст в элементе
                                t.text = ''
                                # Если в r больше нет текста, помечаем для удаления
                                parent = r.getparent()
                                if parent is not None:
                                    # Проверяем, есть ли в r другие непустые w:t
                                    has_text = any(child.text and child.text.strip() for child in r.iter(qn('w:t')))
                                    if not has_text:
                                        # Удаляем пустой r
                                        parent.remove(r)

                        replaced += 1
                        processed += 1
                        break

        return (processed, replaced)

    def get_full_text(self, doc: "DocumentType") -> str:
        """
        Получение полного текста документа включая таблицы.

        Args:
            doc: Объект Document

        Returns:
            Полный текст документа
        """
        from docx.oxml.ns import qn

        text_parts = []

        # Текст из параграфов
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                text_parts.append(paragraph.text)

        # Текст из таблиц через XML (для надёжности)
        for tc in doc._element.body.iter(qn('w:tc')):
            cell_text = ''
            for t in tc.iter(qn('w:t')):
                if t.text:
                    cell_text += t.text
            if cell_text.strip():
                text_parts.append(cell_text)

        return "\n".join(text_parts)
    
    def process_document(self, input_path: str, output_path: str) -> Tuple[bool, str, int]:
        """
        Полная обработка документа.

        Args:
            input_path: Путь к исходному файлу
            output_path: Путь для сохранения результата

        Returns:
            Кортеж (успех, сообщение, количество замен)
        """
        try:
            # Загрузка документа
            doc = self.load_document(input_path)

            # Проверка наличия даты «29» января 2026 г. на первой странице
            has_old_date = self.date_replacer.find_date_in_first_paragraphs(doc)

            if not has_old_date:
                # Даты нет - просто копируем файл без изменений
                self.save_document(doc, output_path)
                return (True, "Файл скопирован без изменений (дата «29» января 2026 г. не найдена)", 0)

            # Обработка параграфов (только первая страница)
            _, para_replaced = self.process_paragraphs(doc, first_page_only=True)

            # Обработка таблиц (только первая страница)
            _, table_replaced = self.process_tables(doc, first_page_only=True)

            total_replaced = para_replaced + table_replaced

            # Сохранение документа
            self.save_document(doc, output_path)

            if total_replaced == 0:
                return (True, "Файл скопирован без изменений (дата не найдена)", 0)

            return (True, f"Заменено дат: {total_replaced}", total_replaced)

        except FileNotFoundError as e:
            return (False, str(e), 0)
        except Exception as e:
            return (False, f"Ошибка обработки: {str(e)}", 0)
    
    def copy_folder_structure(self, src_root: str, dst_root: str) -> None:
        """
        Копирование структуры папок из источника в назначение.
        
        Args:
            src_root: Исходная корневая папка
            dst_root: Целевая корневая папка
        """
        for dirpath, dirnames, filenames in os.walk(src_root):
            # Вычисляем относительный путь
            rel_path = os.path.relpath(dirpath, src_root)
            target_dir = os.path.join(dst_root, rel_path)
            
            # Создаём директорию
            os.makedirs(target_dir, exist_ok=True)
    
    def find_docx_files(self, root_folder: str, exclude_prefix: str = "~$") -> List[str]:
        """
        Поиск всех файлов .docx в папке.
        
        Args:
            root_folder: Корневая папка для поиска
            exclude_prefix: Префикс для исключения файлов
            
        Returns:
            Список путей к файлам .docx
        """
        docx_files = []
        
        for dirpath, dirnames, filenames in os.walk(root_folder):
            for filename in filenames:
                if filename.endswith(".docx") and not filename.startswith(exclude_prefix):
                    filepath = os.path.join(dirpath, filename)
                    docx_files.append(filepath)
        
        return sorted(docx_files)
    
    def get_output_path(self, input_path: str, src_root: str, dst_root: str) -> str:
        """
        Вычисление пути для выходного файла с сохранением структуры.
        
        Args:
            input_path: Путь к исходному файлу
            src_root: Исходная корневая папка
            dst_root: Целевая корневая папка
            
        Returns:
            Путь для выходного файла
        """
        rel_path = os.path.relpath(input_path, src_root)
        return os.path.join(dst_root, rel_path)
