"""Модуль для парсинга Word документа и создания Excel таблицы."""

import argparse
import re
from pathlib import Path

import pandas as pd
from docx import Document


def clean_text(text):
    """Очищает текст от номеров страниц и лишних пробелов."""
    text = re.split(r'…|\.{3,}', text)[0]
    return ' '.join(text.split()).strip()


def parse_content(docx_path):
    """Парсит содержимое Word документа и возвращает данные для Excel."""
    doc = Document(docx_path)
    data = []
    current_id = 1
    current_chapter = None
    current_section = None
    section_pattern = r'^\d+\. '

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        clean_text_value = clean_text(text)

        if clean_text_value.isupper() and len(clean_text_value) > 3:
            current_chapter = {
                'id': current_id,
                'text': clean_text_value,
                'parent': 0,
            }
            data.append(current_chapter)
            current_id += 1
            current_section = None

        elif re.match(section_pattern, clean_text_value):
            if current_chapter is not None:
                current_section = {
                    'id': current_id,
                    'text': clean_text_value,
                    'parent': current_chapter['id'],
                }
                data.append(current_section)
                current_id += 1

        elif current_section is not None:
            data.append(
                {
                    'id': current_id,
                    'text': clean_text_value,
                    'parent': current_section['id'],
                }
            )
            current_id += 1

    return pd.DataFrame(data), current_id - 1


def main():
    """Функция для обработки аргументов и запуска парсинга."""
    parser = argparse.ArgumentParser(
        description='Парсинг Word документа в Excel таблицу.'
    )
    parser.add_argument(
        'filename',
        type=str,
        help='Путь к Word документу (.docx) для обработки',
    )
    args = parser.parse_args()

    input_file = Path(args.filename)
    output_file = Path('output.xlsx')

    try:
        if not input_file.exists():
            raise FileNotFoundError(f'Файл {input_file} не существует!')

        if input_file.suffix.lower() != '.docx':
            raise ValueError(f'Файл {input_file} должен быть в формате .docx!')

        df, num_count = parse_content(input_file)
        df.to_excel(output_file, index=False)

        print(
            'Таблица создана в файле output.xlsx\n'
            f'Добавлено записей: {num_count}'
        )

    except (FileNotFoundError, ValueError) as e:
        print(f'Ошибка: {e}')
    except Exception as e:
        print(f'Непредвиденная ошибка при обработке файла: {e}')


if __name__ == '__main__':
    main()
