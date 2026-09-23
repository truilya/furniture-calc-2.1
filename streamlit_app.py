"""
Приложение: PDF-ТЗ → существующий шаблон Excel.

Установка:
    pip install streamlit pypdf openpyxl pandas requests

Запуск:
    streamlit run app.py

Секреты задаются через переменные окружения или .streamlit/secrets.toml:
    OPENAI_API_KEY = "..."
    YANDEX_API_KEY = "..."
    YANDEX_FOLDER_ID = "..."

Для YandexGPT также поддерживается YANDEX_IAM_TOKEN вместо YANDEX_API_KEY.

Принцип сохранности шаблона:
- существующие строки и листы не вставляются, не удаляются и не перестраиваются;
- формулы, заголовки и служебные листы не перезаписываются;
- запись допускается только в проверенные ячейки товарных блоков;
- если блоков или строк характеристик не хватает, приложение останавливается
  и просит подготовить расширенный шаблон. Это безопаснее, чем сдвигать формулы.

Шаблон-пример содержит старые цены, расчеты и логистические оценки. Приложение
не выдает их за данные нового ТЗ и предупреждает пользователя об их проверке.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
from dataclasses import dataclass
from typing import Any

import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


# ----------------------------- Настройки ----------------------------------

MAX_PDF_BYTES = 25 * 1024 * 1024
MAX_XLSX_BYTES = 25 * 1024 * 1024
MAX_PDF_PAGES = 100
MAX_CHUNK_CHARS = 11_000
REQUEST_TIMEOUT = (10, 120)
MAX_LLM_RESPONSE_TOKENS = 8_000

PRICE_SHEET = "Цены изделий (входящие)"
LOGISTICS_SHEET = "Логистика и пр. расходы"
ECONOMICS_SHEET = "Экономика"

MODEL_CONFIG = {
    "YandexGPT Pro": {
        "provider": "yandex",
        "model": "yandexgpt/latest",
    },
    "YandexGPT Lite": {
        "provider": "yandex",
        "model": "yandexgpt-lite/latest",
    },
    "GPT-4o-mini": {
        "provider": "openai",
        "model": "gpt-4o-mini",
    },
}

# Допускаем только числовой префикс позиции. Это позволяет сохранить различие
# между двумя позициями с одинаковым названием «Диван».
ITEM_NUMBER_RE = re.compile(r"^\s*(\d{1,4})\s*[.)](?:\s|$)")
NUMBERED_PDF_ITEM_RE = re.compile(
    r"(?m)^\s*(\d{1,4})\.\s*"
    r"(?!Общая\b|Стандарт\b|Объем\b|Требования\b|Перечень\b)"
    r"([А-ЯЁ][^\n]{0,140})"
)


class AppError(Exception):
    """Ошибка, которую можно безопасно показать пользователю."""


@dataclass(frozen=True)
class ProcurementItem:
    position_number: int
    object_name: str
    quantity: float
    unit: str
    delivery_address: str | None
    delivery_deadline: str | None
    characteristics: tuple[str, ...]


@dataclass(frozen=True)
class RowBlock:
    position_number: int
    first_row: int
    last_row: int
    existing_name: str


# -------------------------- Секреты и файлы --------------------------------

def get_secret(name: str) -> str:
    """Получить секрет без вывода его значения в интерфейс или журнал."""
    value = os.environ.get(name)
    if value:
        return value
    try:
        return str(st.secrets.get(name, "") or "")
    except (FileNotFoundError, KeyError):
        return ""


def read_uploaded_file(uploaded_file: Any, limit: int, kind: str) -> bytes:
    try:
        data = uploaded_file.getvalue()
    except Exception as exc:
        raise AppError(f"Не удалось прочитать {kind}: {exc}") from exc

    if not data:
        raise AppError(f"{kind} пуст.")
    if len(data) > limit:
        raise AppError(
            f"{kind} превышает допустимый размер "
            f"{limit // (1024 * 1024)} МБ."
        )
    return data


def extract_pdf_text(pdf_bytes: bytes) -> str:
    if not pdf_bytes.startswith(b"%PDF"):
        raise AppError("Загруженный PDF не имеет корректной сигнатуры.")

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
        if reader.is_encrypted:
            raise AppError("Зашифрованный PDF не поддерживается.")
        if len(reader.pages) > MAX_PDF_PAGES:
            raise AppError(f"PDF содержит более {MAX_PDF_PAGES} страниц.")

        pages: list[str] = []
        for index, page in enumerate(reader.pages, start=1):
            text = page.extract_text(extraction_mode="layout") or ""
            if text.strip():
                pages.append(f"\n[СТРАНИЦА {index}]\n{text}")

        result = "\n".join(pages).strip()
    except AppError:
        raise
    except Exception as exc:
        raise AppError(f"Ошибка извлечения текста PDF: {exc}") from exc

    if len(result) < 100:
        raise AppError(
            "В PDF почти нет извлекаемого текста. Вероятно, это скан: "
            "требуется OCR до загрузки файла."
        )
    return result


def open_template(xlsx_bytes: bytes) -> openpyxl.Workbook:
    if not xlsx_bytes.startswith(b"PK"):
        raise AppError("Файл шаблона не похож на корректный XLSX.")

    try:
        return openpyxl.load_workbook(
            io.BytesIO(xlsx_bytes),
            data_only=False,
            read_only=False,
            keep_links=True,
        )
    except Exception as exc:
        raise AppError(f"Не удалось открыть Excel-шаблон: {exc}") from exc


# ---------------------------- LLM API --------------------------------------

SYSTEM_PROMPT = """
Ты извлекаешь структурированные данные из русского технического задания
на закупку. Отвечай ТОЛЬКО JSON-объектом следующего вида:

{
  "procurement_info": {
    "object_of_purchase": "Поставка диванов и кресел",
    "customer": "ГБУЗ Поликлиника Троицкая ДЗМ",
    "total_quantity": 50
  },
  "items": [
    {
      "position_number": "1",
      "name_from_tz": "Секция диванная, деревянный каркас",
      "specifications": [
        "Вид товара: Секция диванная",
        "Тип каркаса: Деревянный",
        "Вид материала обивки: Кожа натуральная",
        "Ширина: Равно 2080 ММ",
        "Глубина: Равно 830 ММ",
        "Высота: Равно 770 ММ"
      ],
      "quantity": 1,
      "unit": "Штука",
      "delivery_logistics": {
        "address": "город Москва, улица 2-я Брестская, дом 6",
        "deadline_days_or_date": "c 1-го по 30-й рабочий день c даты заключения контракта",
        "delivery_form": "В разобранном виде"
      },
      "dimensions_mm": {
        "width_min": 2080,
        "width_max": 2080,
        "depth_min": 830,
        "depth_max": 830,
        "height_min": 770,
        "height_max": 770
      },
      "co_services": [
        "Сборка поставляемого товара в помещениях заказчика",
        "Расстановка поставляемого товара в помещениях заказчика",
        "Погрузочно-разгрузочные работы, включая подъем на этаж"
      ]
    }
  ]
}

Правила:
1. Каждая НУМЕРОВАННАЯ позиция приложения «Перечень объектов закупки» —
   отдельный объект, даже когда названия одинаковы. Не объединяй позиции
   с разными номерами, размерами, материалами или количеством.
2. Извлекай количество и единицу измерения именно товара, а не количество
   ножек, мест или дней. Не придумывай отсутствующие значения.
3. Сохраняй все индивидуальные характеристики позиции, в том числе
   материалы, размеры, диапазоны, отрицания, цвет, особенности конструкции.
   Передавай каждое требование отдельной строкой максимально близко
   к исходной формулировке. Общие юридические разделы не включай.
4. Адрес и срок бери из соответствующей позиции приложения; если они
   указаны только в общих правилах без конкретного значения, верни null.
5. Никаких цен, предполагаемого веса или характеристик «по аналогии».
6. Для фрагмента, содержащего только продолжение позиции, сохраняй ее номер.
   Не добавляй позиции, которых нет в тексте фрагмента.
""".strip()


def llm_post(model_name: str, messages: list[dict[str, str]]) -> dict[str, Any]:
    config = MODEL_CONFIG[model_name]
    provider = config["provider"]

    if provider == "openai":
        api_key = get_secret("OPENAI_API_KEY")
        if not api_key:
            raise AppError("Для GPT-4o-mini задайте OPENAI_API_KEY.")

        url = "https://api.openai.com/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}"}
        payload = {
            "model": config["model"],
            "temperature": 0,
            "max_tokens": MAX_LLM_RESPONSE_TOKENS,
            "response_format": {"type": "json_object"},
            "messages": messages,
        }
    else:
        api_key = get_secret("YANDEX_API_KEY")
        iam_token = get_secret("YANDEX_IAM_TOKEN")
        folder_id = get_secret("YANDEX_FOLDER_ID")
        if not folder_id or not (api_key or iam_token):
            raise AppError(
                "Для YandexGPT задайте YANDEX_FOLDER_ID и "
                "YANDEX_API_KEY либо YANDEX_IAM_TOKEN."
            )

        url = (
            "https://llm.api.cloud.yandex.net/"
            "foundationModels/v1/completion"
        )
        authorization = (
            f"Api-Key {api_key}" if api_key else f"Bearer {iam_token}"
        )
        headers = {"Authorization": authorization}
        payload = {
            "modelUri": f"gpt://{folder_id}/{config['model']}",
            "completionOptions": {
                "stream": False,
                "temperature": 0,
                "maxTokens": str(MAX_LLM_RESPONSE_TOKENS),
            },
            "messages": messages,
            "jsonObject": True,
        }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        body = response.json()

        if provider == "openai":
            content = body["choices"][0]["message"]["content"]
        else:
            content = body["result"]["alternatives"][0]["message"]["text"]

        if not isinstance(content, str):
            raise ValueError("Ответ модели не содержит текстового JSON.")
        parsed = json.loads(content)
        if not isinstance(parsed, dict):
            raise ValueError("Корнем ответа должен быть JSON-объект.")
        return parsed

    except requests.Timeout as exc:
        raise AppError(
            f"Превышено время ожидания API модели {model_name}."
        ) from exc
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response else "неизвестен"
        # Тело ответа не показываем: оно может содержать служебные данные.
        raise AppError(
            f"API модели {model_name} вернул HTTP {status}. "
            "Проверьте доступ, квоты и настройки модели."
        ) from exc
    except requests.RequestException as exc:
        raise AppError(f"Сетевая ошибка API модели {model_name}: {exc}") from exc
    except (KeyError, IndexError, ValueError, json.JSONDecodeError) as exc:
        raise AppError(
            f"Модель {model_name} вернула некорректный структурированный ответ: "
            f"{exc}"
        ) from exc


def split_text(text: str, max_chars: int = MAX_CHUNK_CHARS) -> list[str]:
    """
    Делит текст по страницам/абзацам. Фрагменты перекрываются коротким
    контекстом: это помогает не потерять заголовок позиции на границе.
    """
    if len(text) <= max_chars:
        return [text]

    paragraphs = re.split(r"\n\s*\n", text)
    chunks: list[str] = []
    current = ""

    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if not paragraph:
            continue

        if len(paragraph) > max_chars:
            if current:
                chunks.append(current)
                current = ""
            for start in range(0, len(paragraph), max_chars - 500):
                chunks.append(paragraph[start:start + max_chars])
            continue

        proposed = f"{current}\n\n{paragraph}" if current else paragraph
        if len(proposed) > max_chars and current:
            chunks.append(current)
            current = f"[КОНТЕКСТ ПРЕДЫДУЩЕГО ФРАГМЕНТА]\n{current[-500:]}\n\n{paragraph}"
        else:
            current = proposed

    if current:
        chunks.append(current)
    return chunks


def clean_text(value: Any, field: str, nullable: bool = False) -> str | None:
    if value is None and nullable:
        return None
    if not isinstance(value, str):
        raise AppError(f"Поле «{field}» имеет неверный тип.")
    result = re.sub(r"\s+", " ", value).strip()
    if not result and not nullable:
        raise AppError(f"Обязательное поле «{field}» пусто.")
    return result or None


def validate_llm_item(raw: Any) -> ProcurementItem:
    if not isinstance(raw, dict):
        raise AppError("Одна из позиций LLM не является JSON-объектом.")

    number = raw.get("position_number")
    if isinstance(number, bool) or not isinstance(number, int) or number < 1:
        raise AppError("У позиции отсутствует корректный position_number.")

    quantity = raw.get("quantity")
    if (
        isinstance(quantity, bool)
        or not isinstance(quantity, (int, float))
        or not 0 < float(quantity) < 1_000_000_000
    ):
        raise AppError(f"Позиция №{number}: некорректное количество.")

    characteristics = raw.get("characteristics")
    if not isinstance(characteristics, list) or not characteristics:
        raise AppError(f"Позиция №{number}: отсутствуют характеристики.")

    cleaned: list[str] = []
    seen: set[str] = set()
    for value in characteristics:
        line = clean_text(value, f"характеристика позиции №{number}")
        assert line is not None
        key = line.casefold()
        if key not in seen:
            seen.add(key)
            cleaned.append(line)

    name = clean_text(raw.get("object_name"), "object_name")
    unit = clean_text(raw.get("unit"), "unit")
    assert name is not None and unit is not None

    return ProcurementItem(
        position_number=number,
        object_name=name,
        quantity=float(quantity),
        unit=unit,
        delivery_address=clean_text(
            raw.get("delivery_address"),
            "delivery_address",
            nullable=True,
        ),
        delivery_deadline=clean_text(
            raw.get("delivery_deadline"),
            "delivery_deadline",
            nullable=True,
        ),
        characteristics=tuple(cleaned),
    )


def extract_items_with_llm(text: str, model_name: str) -> list[ProcurementItem]:
    chunks = split_text(text)
    merged: dict[int, ProcurementItem] = {}

    for index, chunk in enumerate(chunks, start=1):
        result = llm_post(
            model_name,
            [
                {"role": "system", "text": SYSTEM_PROMPT}
                if MODEL_CONFIG[model_name]["provider"] == "yandex"
                else {"role": "system", "content": SYSTEM_PROMPT},
                {
                    "role": "user",
                    **(
                        {"text": f"Фрагмент {index}/{len(chunks)}:\n\n{chunk}"}
                        if MODEL_CONFIG[model_name]["provider"] == "yandex"
                        else {
                            "content": (
                                f"Фрагмент {index}/{len(chunks)}:\n\n{chunk}"
                            )
                        }
                    ),
                },
            ],
        )

        raw_items = result.get("items")
        if not isinstance(raw_items, list):
            raise AppError(
                f"Фрагмент {index}: ответ модели не содержит массив items."
            )

        for raw in raw_items:
            item = validate_llm_item(raw)
            previous = merged.get(item.position_number)

            if previous is None:
                merged[item.position_number] = item
                continue

            # Повтор позиции возможен на границе фрагментов. Разные количества
            # либо названия — повод остановиться, а не молча выбрать одно.
            if (
                previous.object_name.casefold() != item.object_name.casefold()
                or previous.quantity != item.quantity
                or previous.unit.casefold() != item.unit.casefold()
            ):
                raise AppError(
                    f"Противоречивые данные LLM для позиции "
                    f"№{item.position_number}. Проверьте PDF."
                )

            combined = list(previous.characteristics)
            existing = {line.casefold() for line in combined}
            for line in item.characteristics:
                if line.casefold() not in existing:
                    combined.append(line)
                    existing.add(line.casefold())

            merged[item.position_number] = ProcurementItem(
                position_number=item.position_number,
                object_name=previous.object_name,
                quantity=previous.quantity,
                unit=previous.unit,
                delivery_address=(
                    previous.delivery_address or item.delivery_address
                ),
                delivery_deadline=(
                    previous.delivery_deadline or item.delivery_deadline
                ),
                characteristics=tuple(combined),
            )

    if not merged:
        raise AppError("Модель не нашла товарных позиций в PDF.")

    return [merged[number] for number in sorted(merged)]


# ------------------------- Проверка шаблона --------------------------------

def item_number(value: Any) -> int | None:
    if not isinstance(value, str):
        return None
    match = ITEM_NUMBER_RE.match(value)
    return int(match.group(1)) if match else None


def find_header_row(ws: openpyxl.worksheet.worksheet.Worksheet) -> int:
    """Находит строку заголовка конкретной таблицы, не предполагая ее номер."""
    for row in ws.iter_rows(
        min_row=1,
        max_row=min(ws.max_row, 30),
        max_col=min(ws.max_column, 7),
    ):
        values = [
            str(cell.value or "").replace("\n", " ").casefold()
            for cell in row
        ]
        if any("наименование товара из тз" in value for value in values):
            return row[0].row

    raise AppError(
        f"Лист «{ws.title}»: не найден ожидаемый заголовок "
        "«Наименование товара из ТЗ»."
    )


def detect_blocks(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    header_row: int,
) -> list[RowBlock]:
    """
    Границы блока — следующая пронумерованная позиция либо строка итогов.
    Строки заголовков и итогов не попадают в доступный диапазон.
    """
    starts: list[tuple[int, int, str]] = []
    stop_row = ws.max_row + 1

    for row in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(row, 1).value
        number = item_number(value)
        if number is not None:
            original = str(value)
            name = re.sub(
                r"^\s*\d{1,4}\s*[.)]\s*",
                "",
                original,
                count=1,
            ).strip()
            starts.append((row, number, name))
        elif starts and isinstance(value, str):
            lowered = value.strip().casefold()
            if lowered.startswith(("итого", "итоговая", "всего")):
                stop_row = row
                break

    if not starts:
        raise AppError(
            f"Лист «{ws.title}»: не найдены нумерованные товарные блоки."
        )

    numbers = [number for _, number, _ in starts]
    if len(numbers) != len(set(numbers)):
        raise AppError(
            f"Лист «{ws.title}»: повторяются номера товарных блоков."
        )

    return [
        RowBlock(
            position_number=number,
            first_row=row,
            last_row=(
                starts[index + 1][0] - 1
                if index + 1 < len(starts)
                else stop_row - 1
            ),
            existing_name=name,
        )
        for index, (row, number, name) in enumerate(starts)
    ]


def assert_editable(ws: Any, row: int, column: int) -> None:
    cell = ws.cell(row, column)
    if isinstance(cell, MergedCell):
        raise AppError(
            f"Нельзя записать в объединенную ячейку "
            f"{ws.title}!{get_column_letter(column)}{row}."
        )
    if cell.data_type == "f" or (
        isinstance(cell.value, str) and cell.value.startswith("=")
    ):
        raise AppError(
            f"Ячейка {ws.title}!{cell.coordinate} содержит формулу. "
            "Изменения отменены."
        )


def set_cell(
    ws: Any,
    row: int,
    column: int,
    value: str | float | int | None,
) -> None:
    assert_editable(ws, row, column)
    ws.cell(row, column).value = value


def assert_has_capacity(
    ws: Any,
    block: RowBlock,
    required_lines: int,
) -> None:
    available = block.last_row - block.first_row + 1
    if required_lines > available:
        raise AppError(
            f"Лист «{ws.title}», позиция №{block.position_number}: "
            f"нужно {required_lines} строк характеристик, в шаблоне "
            f"доступно {available}. Подготовьте шаблон с большим блоком; "
            "автоматическая вставка строк отключена для защиты формул."
        )


def validate_mapping(
    workbook: openpyxl.Workbook,
    items: list[ProcurementItem],
) -> dict[str, dict[int, RowBlock]]:
    needed = {PRICE_SHEET, LOGISTICS_SHEET, ECONOMICS_SHEET}
    missing_sheets = needed.difference(workbook.sheetnames)
    if missing_sheets:
        raise AppError(
            "В шаблоне отсутствуют листы: "
            + ", ".join(sorted(missing_sheets))
        )

    mappings: dict[str, dict[int, RowBlock]] = {}
    for sheet_name in (PRICE_SHEET, LOGISTICS_SHEET, ECONOMICS_SHEET):
        ws = workbook[sheet_name]
        blocks = detect_blocks(ws, find_header_row(ws))
        mappings[sheet_name] = {
            block.position_number: block for block in blocks
        }

    pdf_numbers = {item.position_number for item in items}
    for sheet_name, mapping in mappings.items():
        template_numbers = set(mapping)
        if pdf_numbers != template_numbers:
            missing = sorted(pdf_numbers - template_numbers)
            surplus = sorted(template_numbers - pdf_numbers)
            raise AppError(
                f"Лист «{sheet_name}»: позиции PDF и шаблона не совпадают. "
                f"Нет блоков для {missing}; лишние блоки {surplus}. "
                "Шаблон с другими закупочными позициями нужно сначала "
                "адаптировать. Подмена по порядку небезопасна."
            )

    return mappings


def dimensions_only(characteristics: tuple[str, ...]) -> list[str]:
    """
    Лист логистики содержит только строки габаритов; количество мест,
    масса нагрузки и размеры упаковки не подменяют размеры изделия.
    """
    prefixes = (
        "высота",
        "глубина",
        "ширина",
        "длина",
        "толщина",
        "диаметр",
    )
    return [
        line for line in characteristics
        if line.casefold().startswith(prefixes)
    ]


def write_characteristics(
    ws: Any,
    block: RowBlock,
    column: int,
    values: list[str],
) -> None:
    assert_has_capacity(ws, block, len(values))

    # Очистка остатка старого блока выполняется только после проверки
    # каждой ячейки. Формулы, если они здесь есть, вызывают отказ.
    for row in range(block.first_row, block.last_row + 1):
        assert_editable(ws, row, column)

    for offset, line in enumerate(values):
        set_cell(ws, block.first_row + offset, column, line)
    for row in range(block.first_row + len(values), block.last_row + 1):
        set_cell(ws, row, column, None)


def populate_workbook(
    workbook: openpyxl.Workbook,
    items: list[ProcurementItem],
) -> None:
    mapping = validate_mapping(workbook, items)
    prices = workbook[PRICE_SHEET]
    logistics = workbook[LOGISTICS_SHEET]
    economics = workbook[ECONOMICS_SHEET]

    # Лист «Доп.информация» примера содержит текст ДРУГОЙ закупки;
    # служебные листы — старые габариты и оценки. Не меняем их автоматически.
    # Пользователю необходимо проверить эти листы отдельно.

    for item in items:
        price_block = mapping[PRICE_SHEET][item.position_number]
        logistics_block = mapping[LOGISTICS_SHEET][item.position_number]
        economics_block = mapping[ECONOMICS_SHEET][item.position_number]

        price_name = (
            f"{item.position_number}.\n{item.object_name}"
        )
        quantity: int | float = (
            int(item.quantity)
            if item.quantity.is_integer()
            else item.quantity
        )

        # «Цены изделий (входящие)»:
        # A — номер и название; D — характеристики; E — количество.
        # B, C, F:I оставляем нетронутыми: поставщик, изображения, цены
        # и вычисляемые суммы не извлекаются из ТЗ.
        set_cell(prices, price_block.first_row, 1, price_name)
        set_cell(prices, price_block.first_row, 5, quantity)
        write_characteristics(
            prices,
            price_block,
            4,
            list(item.characteristics),
        )

        # «Логистика и пр. расходы»:
        # A — позиция; C — размерные характеристики; D — количество;
        # E — единица. Расчетные объемы и массы не трогаем.
        set_cell(logistics, logistics_block.first_row, 1, price_name)
        set_cell(logistics, logistics_block.first_row, 4, quantity)
        set_cell(logistics, logistics_block.first_row, 5, item.unit)
        write_characteristics(
            logistics,
            logistics_block,
            3,
            dimensions_only(item.characteristics),
        )

        # «Экономика»: обновляем только идентификатор и количество.
        # Цены, прибыль, источники товара и их формулы сохраняются.
        set_cell(economics, economics_block.first_row, 1, price_name)
        set_cell(economics, economics_block.first_row, 3, quantity)

    # Excel пересчитает сохраненные формулы при открытии. openpyxl
    # вычислять формулы не умеет.
    try:
        workbook.calculation = openpyxl.workbook.properties.CalcProperties(
            calcMode="auto",
            fullCalcOnLoad=True,
            forceFullCalc=True,
        )
    except AttributeError:
        # Для версий openpyxl с другим интерфейсом расчета не меняем
        # настройки книги: сами формулы остаются сохраненными.
        pass


def workbook_to_bytes(workbook: openpyxl.Workbook) -> bytes:
    try:
        output = io.BytesIO()
        workbook.save(output)
        result = output.getvalue()
        if not result:
            raise ValueError("Получен пустой XLSX.")
        return result
    except Exception as exc:
        raise AppError(f"Не удалось сохранить заполненный XLSX: {exc}") from exc


def items_to_dataframe(items: list[ProcurementItem]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "№": item.position_number,
                "Объект закупки": item.object_name,
                "Количество": item.quantity,
                "Ед. изм.": item.unit,
                "Адрес поставки": item.delivery_address or "Не указан",
                "Срок поставки": item.delivery_deadline or "Не указан",
                "Характеристик": len(item.characteristics),
            }
            for item in items
        ]
    )


# ------------------------------ Streamlit ----------------------------------

def main() -> None:
    st.set_page_config(
        page_title="Техническое задание → Excel",
        page_icon="📋",
        layout="wide",
    )
    st.title("Техническое задание → Excel-шаблон")
    st.write(
        "Загрузите PDF с текстовым слоем и подготовленный XLSX-шаблон. "
        "Товары сопоставляются по номерам позиций, а не по совпадению названий."
    )

    model_name = st.sidebar.selectbox(
        "Модель для обработки текста",
        list(MODEL_CONFIG),
    )
    st.sidebar.caption(
        "Ключи API задаются через переменные окружения или "
        ".streamlit/secrets.toml."
    )

    st.subheader("Исходные файлы")
    pdf_upload = st.file_uploader(
        "Техническое задание, PDF",
        type=["pdf"],
        key="pdf_upload",
    )
    xlsx_upload = st.file_uploader(
        "Предварительно подготовленный шаблон, XLSX",
        type=["xlsx"],
        key="xlsx_upload",
    )

    if not pdf_upload or not xlsx_upload:
        return

    # Хеш входов предотвращает выдачу результата, созданного для других
    # файлов или другой модели, после изменения загрузки/выбора.
    try:
        pdf_bytes = read_uploaded_file(
            pdf_upload, MAX_PDF_BYTES, "PDF-файл"
        )
        xlsx_bytes = read_uploaded_file(
            xlsx_upload, MAX_XLSX_BYTES, "Excel-шаблон"
        )
    except AppError as exc:
        st.error(str(exc))
        return

    fingerprint = hashlib.sha256(
        pdf_bytes + b"\0" + xlsx_bytes + b"\0" + model_name.encode()
    ).hexdigest()

    if st.session_state.get("result_fingerprint") != fingerprint:
        st.session_state.pop("result_binary", None)
        st.session_state.pop("result_dataframe", None)
        st.session_state.pop("result_fingerprint", None)

    if st.button(
        "Извлечь данные и заполнить шаблон",
        type="primary",
        disabled=bool(st.session_state.get("result_binary")),
    ):
        with st.spinner("Чтение PDF, обработка моделью и проверка шаблона…"):
            try:
                text = extract_pdf_text(pdf_bytes)
                items = extract_items_with_llm(text, model_name)
                workbook = open_template(xlsx_bytes)
                populate_workbook(workbook, items)
                result_binary = workbook_to_bytes(workbook)

                st.session_state["result_binary"] = result_binary
                st.session_state["result_dataframe"] = items_to_dataframe(
                    items
                )
                st.session_state["result_fingerprint"] = fingerprint

            except AppError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(
                    "Непредвиденная ошибка обработки. "
                    f"Подробности: {type(exc).__name__}: {exc}"
                )

    if st.session_state.get("result_binary"):
        st.success("Файл подготовлен.")
        st.subheader("Извлеченные позиции")
        st.dataframe(
            st.session_state["result_dataframe"],
            use_container_width=True,
            hide_index=True,
        )

        st.warning(
            "Проверьте результат перед использованием: существующие цены, "
            "массы, объемы, тексты листа «Доп.информация» и служебные "
            "расчеты шаблона могли относиться к другой закупке. "
            "Приложение намеренно не заменяет их предположениями ИИ. "
            "Формулы пересчитываются Excel при открытии файла."
        )

        st.download_button(
            label="Скачать заполненный XLSX",
            data=st.session_state["result_binary"],
            file_name="ТЗ_заполненный_шаблон.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_result",
        )


if __name__ == "__main__":
    main()