"""
PDF-ТЗ → подготовленный Excel-шаблон.

Установка:
    pip install streamlit pypdf openpyxl pandas requests

Запуск:
    streamlit run app.py

Безопасность и ограничения:
- API-ключ вводится пользователем в боковой панели и не записывается в XLSX.
- Для публичных провайдеров используется HTTPS.
- Собственный endpoint допускает HTTP для локального Ollama/vLLM.
- PDF должен содержать извлекаемый текст: OCR сканов здесь не выполняется.
- Приложение не вставляет и не удаляет строки шаблона, чтобы не сдвигать
  формулы и связанные диапазоны.
- Товарные блоки сопоставляются строго по номерам позиций.
- Если PDF и шаблон относятся к разным закупкам, обработка останавливается.
- Существующие цены, логистические оценки и текст служебных листов не
  объявляются данными нового ТЗ и не заменяются догадками модели.
"""

from __future__ import annotations

import hashlib
import io
import json
import re
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlparse

import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.cell.cell import MergedCell
from pypdf import PdfReader


# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------

PROVIDERS: dict[str, dict[str, Any]] = {
    "OpenAI": {
        "base_url": "https://api.openai.com/v1",
        "models": [
            "gpt-4o-mini",
            "gpt-4o",
            "gpt-5.6-sol",
            "gpt-5.6-luna",
        ],
    },
    "OpenRouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "models": [
            "anthropic/claude-sonnet-5",
            "openai/gpt-5.6-sol",
            "google/gemini-3.1-flash-lite",
            "deepseek/deepseek-v4-pro",
            "qwen/qwen-2.5-vl-72b-instruct",
            "openai/gpt-4o-mini",
            "anthropic/claude-3.5-haiku",
            "meta-llama/llama-3.1-70b-instruct",
        ],
    },
    "Groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "models": [
            "llama-3.3-70b-versatile",
            "llama-3.1-8b-instant",
            "mixtral-8x7b-32768",
        ],
    },
    "Свой endpoint": {
        "base_url": "http://localhost:11434/v1",
        "models": [
            "deepseek-r1:70b",
            "qwen2.5-vl:72b",
        ],
    },
}

PRICE_SHEET = "Цены изделий (входящие)"
LOGISTICS_SHEET = "Логистика и пр. расходы"
ECONOMICS_SHEET = "Экономика"

MAX_PDF_BYTES = 25 * 1024 * 1024
MAX_XLSX_BYTES = 25 * 1024 * 1024
MAX_PDF_PAGES = 100

# Размер ограничивает один запрос. Длинные документы обрабатываются частями.
MAX_CHUNK_CHARS = 10_000
MAX_OUTPUT_TOKENS = 8_000

REQUEST_TIMEOUT = (10, 180)

ITEM_NUMBER_RE = re.compile(r"^\s*(\d{1,4})\s*[.)](?:\s|$)")


class AppError(Exception):
    """Ошибка, сообщение которой можно показать в интерфейсе."""


@dataclass(frozen=True)
class LLMSettings:
    provider: str
    model: str
    base_url: str
    api_key: str


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


SYSTEM_PROMPT = """
Ты — интеллектуальный модуль парсинга государственных закупок (ТЗ) на поставку мебели.
Твоя задача — извлечь данные из текста ТЗ и вернуть строго валидный JSON:

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
1. Каждая нумерованная товарная позиция приложения «Перечень объектов
   закупки» — отдельный объект. Не объединяй одинаковые названия.
2. Сохраняй номер позиции из документа. Не перенумеровывай товары.
3. Количество и единицу измерения бери из строки объема товара, а не
   из характеристик вроде количества мест или ножек.
4. Сохраняй индивидуальные характеристики позиции отдельными строками,
   включая материалы, габариты, диапазоны, цвет, отрицания и конструкцию.
5. Общие юридические разделы не включай в характеристики товара.
6. Адрес и срок бери из соответствующей позиции; если точного значения
   нет, верни null. Ничего не придумывай.
7. Цены, предполагаемую массу и параметры «по аналогии» не добавляй.
8. Если фрагмент содержит только продолжение позиции, используй ее
   действительный номер. Не добавляй позиции вне этого фрагмента.
""".strip()


# ---------------------------------------------------------------------------
# Чтение файлов
# ---------------------------------------------------------------------------

def read_uploaded_file(uploaded_file: Any, limit: int, label: str) -> bytes:
    try:
        data = uploaded_file.getvalue()
    except Exception as exc:
        raise AppError(f"Не удалось прочитать {label}: {exc}") from exc

    if not data:
        raise AppError(f"{label} пуст.")

    if len(data) > limit:
        raise AppError(
            f"{label} превышает допустимый размер "
            f"{limit // (1024 * 1024)} МБ."
        )

    return data


def extract_pdf_text(pdf_bytes: bytes) -> str:
    if not pdf_bytes.startswith(b"%PDF"):
        raise AppError("Файл не имеет корректной сигнатуры PDF.")

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)

        if reader.is_encrypted:
            raise AppError("Зашифрованный PDF не поддерживается.")

        if len(reader.pages) > MAX_PDF_PAGES:
            raise AppError(
                f"В PDF более {MAX_PDF_PAGES} страниц."
            )

        pages: list[str] = []

        for page_number, page in enumerate(reader.pages, start=1):
            text = page.extract_text(extraction_mode="layout") or ""

            if text.strip():
                pages.append(
                    f"\n[СТРАНИЦА {page_number}]\n{text}"
                )

        result = "\n".join(pages).strip()

    except AppError:
        raise
    except Exception as exc:
        raise AppError(
            f"Не удалось извлечь текст из PDF: {exc}"
        ) from exc

    if len(result) < 100:
        raise AppError(
            "В PDF почти нет извлекаемого текста. "
            "Для сканированного документа сначала выполните OCR."
        )

    return result


def open_template(xlsx_bytes: bytes) -> openpyxl.Workbook:
    if not xlsx_bytes.startswith(b"PK"):
        raise AppError("Загруженный файл не похож на XLSX.")

    try:
        return openpyxl.load_workbook(
            io.BytesIO(xlsx_bytes),
            data_only=False,
            read_only=False,
            keep_links=True,
        )
    except Exception as exc:
        raise AppError(
            f"Не удалось открыть Excel-шаблон: {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# Настройки API и вызов модели
# ---------------------------------------------------------------------------

def normalize_base_url(provider: str, raw_url: str) -> str:
    url = raw_url.strip().rstrip("/")

    if not url:
        raise AppError("Адрес API endpoint не указан.")

    parsed = urlparse(url)

    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        raise AppError(
            "Endpoint должен быть полным URL, например "
            "http://localhost:11434/v1."
        )

    if provider != "Свой endpoint" and parsed.scheme != "https":
        raise AppError(
            "Для публичного API требуется HTTPS."
        )

    if parsed.username or parsed.password:
        raise AppError(
            "Не помещайте учетные данные в URL endpoint."
        )

    return url


def call_llm(
    settings: LLMSettings,
    user_prompt: str,
) -> dict[str, Any]:
    """
    Все четыре варианта используют Chat Completions-совместимый endpoint.

    response_format=json_object намеренно не включен: поддержка этого
    параметра отличается у моделей и прокси. JSON строго запрашивается
    промптом, затем проверяется на стороне приложения.
    """
    url = f"{settings.base_url}/chat/completions"

    headers = {
        "Content-Type": "application/json",
    }

    if settings.api_key:
        headers["Authorization"] = f"Bearer {settings.api_key}"

    if settings.provider == "OpenRouter":
        headers["HTTP-Referer"] = "http://localhost:8501"
        headers["X-Title"] = "Procurement TS to Excel"

    payload = {
        "model": settings.model,
        "temperature": 0,
        "max_tokens": MAX_OUTPUT_TOKENS,
        "messages": [
            {
                "role": "system",
                "content": SYSTEM_PROMPT,
            },
            {
                "role": "user",
                "content": user_prompt,
            },
        ],
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

        content = body["choices"][0]["message"]["content"]

        if isinstance(content, list):
            # Некоторые совместимые API отдают текст отдельными частями.
            content = "".join(
                part.get("text", "")
                for part in content
                if isinstance(part, dict)
            )

        if not isinstance(content, str) or not content.strip():
            raise ValueError("В ответе модели отсутствует текст.")

        return parse_json_response(content)

    except requests.Timeout as exc:
        raise AppError(
            f"Превышено время ожидания API: "
            f"{settings.provider} / {settings.model}."
        ) from exc

    except requests.HTTPError as exc:
        status = (
            exc.response.status_code
            if exc.response is not None
            else "неизвестен"
        )
        raise AppError(
            f"API {settings.provider} вернул HTTP {status} для модели "
            f"«{settings.model}». Проверьте API-ключ, доступность "
            "идентификатора модели, лимиты и URL endpoint."
        ) from exc

    except requests.RequestException as exc:
        raise AppError(
            f"Ошибка подключения к {settings.provider}: {exc}"
        ) from exc

    except (KeyError, IndexError, TypeError, ValueError) as exc:
        raise AppError(
            f"Модель «{settings.model}» вернула некорректный ответ: "
            f"{exc}"
        ) from exc


def parse_json_response(content: str) -> dict[str, Any]:
    text = content.strip()

    # Допускаем только распространенную обертку ```json ... ```.
    # Произвольный текст вокруг JSON не принимаем.
    fenced = re.fullmatch(
        r"```(?:json)?\s*(\{.*\})\s*```",
        text,
        flags=re.DOTALL | re.IGNORECASE,
    )

    if fenced:
        text = fenced.group(1)

    parsed = json.loads(text)

    if not isinstance(parsed, dict):
        raise ValueError("Корнем ответа должен быть JSON-объект.")

    return parsed


def split_text(
    text: str,
    max_chars: int = MAX_CHUNK_CHARS,
) -> list[str]:
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

            step = max_chars - 500

            for start in range(0, len(paragraph), step):
                chunks.append(
                    paragraph[start:start + max_chars]
                )

            continue

        candidate = (
            f"{current}\n\n{paragraph}"
            if current
            else paragraph
        )

        if len(candidate) > max_chars and current:
            chunks.append(current)

            context = current[-500:]

            current = (
                "[КОНТЕКСТ ПРЕДЫДУЩЕГО ФРАГМЕНТА]\n"
                f"{context}\n\n{paragraph}"
            )
        else:
            current = candidate

    if current:
        chunks.append(current)

    return chunks


def clean_text(
    value: Any,
    field: str,
    nullable: bool = False,
) -> str | None:
    if value is None and nullable:
        return None

    if not isinstance(value, str):
        raise AppError(
            f"Поле «{field}» имеет неверный тип."
        )

    result = re.sub(r"\s+", " ", value).strip()

    if not result and not nullable:
        raise AppError(
            f"Обязательное поле «{field}» пусто."
        )

    return result or None


def validate_llm_item(raw: Any) -> ProcurementItem:
    if not isinstance(raw, dict):
        raise AppError(
            "Одна из позиций ответа модели не является JSON-объектом."
        )

    number = raw.get("position_number")

    if (
        isinstance(number, bool)
        or not isinstance(number, int)
        or number < 1
    ):
        raise AppError(
            "Некорректный номер одной из позиций."
        )

    quantity = raw.get("quantity")

    if (
        isinstance(quantity, bool)
        or not isinstance(quantity, (int, float))
        or not 0 < float(quantity) < 1_000_000_000
    ):
        raise AppError(
            f"Позиция №{number}: некорректное количество."
        )

    characteristics = raw.get("characteristics")

    if (
        not isinstance(characteristics, list)
        or not characteristics
    ):
        raise AppError(
            f"Позиция №{number}: отсутствуют характеристики."
        )

    unique_lines: list[str] = []
    seen: set[str] = set()

    for raw_line in characteristics:
        line = clean_text(
            raw_line,
            f"характеристика позиции №{number}",
        )

        assert line is not None

        normalized = line.casefold()

        if normalized not in seen:
            seen.add(normalized)
            unique_lines.append(line)

    name = clean_text(
        raw.get("object_name"),
        "object_name",
    )
    unit = clean_text(
        raw.get("unit"),
        "unit",
    )

    assert name is not None
    assert unit is not None

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
        characteristics=tuple(unique_lines),
    )


def extract_items_with_llm(
    text: str,
    settings: LLMSettings,
) -> list[ProcurementItem]:
    chunks = split_text(text)
    merged: dict[int, ProcurementItem] = {}

    for index, chunk in enumerate(chunks, start=1):
        result = call_llm(
            settings,
            (
                f"Фрагмент документа {index}/{len(chunks)}.\n"
                "Извлеки позиции только из этого фрагмента.\n\n"
                f"{chunk}"
            ),
        )

        raw_items = result.get("items")

        if not isinstance(raw_items, list):
            raise AppError(
                f"Фрагмент {index}: ответ модели "
                "не содержит массив items."
            )

        for raw_item in raw_items:
            item = validate_llm_item(raw_item)
            previous = merged.get(item.position_number)

            if previous is None:
                merged[item.position_number] = item
                continue

            if (
                previous.object_name.casefold()
                != item.object_name.casefold()
                or previous.quantity != item.quantity
                or previous.unit.casefold()
                != item.unit.casefold()
            ):
                raise AppError(
                    "Модель вернула противоречивые данные "
                    f"для позиции №{item.position_number}."
                )

            combined = list(previous.characteristics)
            existing = {
                line.casefold()
                for line in combined
            }

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
                    previous.delivery_address
                    or item.delivery_address
                ),
                delivery_deadline=(
                    previous.delivery_deadline
                    or item.delivery_deadline
                ),
                characteristics=tuple(combined),
            )

    if not merged:
        raise AppError(
            "Модель не обнаружила товарных позиций в PDF."
        )

    return [
        merged[number]
        for number in sorted(merged)
    ]


# ---------------------------------------------------------------------------
# Работа с Excel
# ---------------------------------------------------------------------------

def item_number(value: Any) -> int | None:
    if not isinstance(value, str):
        return None

    match = ITEM_NUMBER_RE.match(value)

    if not match:
        return None

    return int(match.group(1))


def find_header_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
) -> int:
    for row in ws.iter_rows(
        min_row=1,
        max_row=min(ws.max_row, 30),
        max_col=min(ws.max_column, 7),
    ):
        values = [
            str(cell.value or "")
            .replace("\n", " ")
            .casefold()
            for cell in row
        ]

        if any(
            "наименование товара из тз" in value
            for value in values
        ):
            return row[0].row

    raise AppError(
        f"Лист «{ws.title}»: не найден заголовок "
        "«Наименование товара из ТЗ»."
    )


def detect_blocks(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    header_row: int,
) -> list[RowBlock]:
    starts: list[tuple[int, int]] = []
    stop_row = ws.max_row + 1

    for row in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(row, 1).value
        number = item_number(value)

        if number is not None:
            starts.append((row, number))
            continue

        if starts and isinstance(value, str):
            if value.strip().casefold().startswith(
                ("итого", "итоговая", "всего")
            ):
                stop_row = row
                break

    if not starts:
        raise AppError(
            f"Лист «{ws.title}»: "
            "не найдены пронумерованные товарные блоки."
        )

    numbers = [
        number
        for _, number in starts
    ]

    if len(numbers) != len(set(numbers)):
        raise AppError(
            f"Лист «{ws.title}»: "
            "номера товарных блоков повторяются."
        )

    blocks: list[RowBlock] = []

    for index, (first_row, number) in enumerate(starts):
        last_row = (
            starts[index + 1][0] - 1
            if index + 1 < len(starts)
            else stop_row - 1
        )

        blocks.append(
            RowBlock(
                position_number=number,
                first_row=first_row,
                last_row=last_row,
            )
        )

    return blocks


def assert_editable(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    column: int,
) -> None:
    cell = ws.cell(row, column)

    if isinstance(cell, MergedCell):
        raise AppError(
            f"Нельзя записать в объединенную ячейку "
            f"«{ws.title}!{cell.coordinate}»."
        )

    if (
        cell.data_type == "f"
        or (
            isinstance(cell.value, str)
            and cell.value.startswith("=")
        )
    ):
        raise AppError(
            f"Ячейка «{ws.title}!{cell.coordinate}» "
            "содержит формулу."
        )


def set_cell(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    column: int,
    value: str | int | float | None,
) -> None:
    assert_editable(ws, row, column)
    ws.cell(row, column).value = value


def check_capacity(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    block: RowBlock,
    needed: int,
) -> None:
    capacity = (
        block.last_row
        - block.first_row
        + 1
    )

    if needed > capacity:
        raise AppError(
            f"Лист «{ws.title}», позиция "
            f"№{block.position_number}: "
            f"требуется {needed} строк, доступно {capacity}. "
            "Подготовьте более вместительный шаблон. "
            "Вставка строк отключена для защиты формул."
        )


def write_characteristics(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    block: RowBlock,
    column: int,
    values: list[str],
) -> None:
    check_capacity(
        ws,
        block,
        len(values),
    )

    for row in range(
        block.first_row,
        block.last_row + 1,
    ):
        assert_editable(
            ws,
            row,
            column,
        )

    for index, value in enumerate(values):
        set_cell(
            ws,
            block.first_row + index,
            column,
            value,
        )

    for row in range(
        block.first_row + len(values),
        block.last_row + 1,
    ):
        set_cell(
            ws,
            row,
            column,
            None,
        )


def dimension_characteristics(
    characteristics: tuple[str, ...],
) -> list[str]:
    prefixes = (
        "высота",
        "глубина",
        "ширина",
        "длина",
        "толщина",
        "диаметр",
    )

    return [
        line
        for line in characteristics
        if line.casefold().startswith(prefixes)
    ]


def validate_template_mapping(
    workbook: openpyxl.Workbook,
    items: list[ProcurementItem],
) -> dict[str, dict[int, RowBlock]]:
    required_sheets = {
        PRICE_SHEET,
        LOGISTICS_SHEET,
        ECONOMICS_SHEET,
    }

    missing_sheets = (
        required_sheets
        - set(workbook.sheetnames)
    )

    if missing_sheets:
        raise AppError(
            "В шаблоне отсутствуют листы: "
            + ", ".join(sorted(missing_sheets))
        )

    result: dict[str, dict[int, RowBlock]] = {}

    for sheet_name in (
        PRICE_SHEET,
        LOGISTICS_SHEET,
        ECONOMICS_SHEET,
    ):
        ws = workbook[sheet_name]
        header_row = find_header_row(ws)
        blocks = detect_blocks(ws, header_row)

        result[sheet_name] = {
            block.position_number: block
            for block in blocks
        }

    pdf_numbers = {
        item.position_number
        for item in items
    }

    for sheet_name, mapping in result.items():
        template_numbers = set(mapping)

        if template_numbers != pdf_numbers:
            missing = sorted(
                pdf_numbers - template_numbers
            )
            surplus = sorted(
                template_numbers - pdf_numbers
            )

            raise AppError(
                f"Лист «{sheet_name}» не соответствует PDF. "
                f"Отсутствующие блоки: {missing}; "
                f"лишние блоки: {surplus}. "
                "Автоматическая подмена позиций "
                "другой закупки запрещена."
            )

    return result


def populate_workbook(
    workbook: openpyxl.Workbook,
    items: list[ProcurementItem],
) -> None:
    mapping = validate_template_mapping(
        workbook,
        items,
    )

    prices = workbook[PRICE_SHEET]
    logistics = workbook[LOGISTICS_SHEET]
    economics = workbook[ECONOMICS_SHEET]

    for item in items:
        price_block = mapping[PRICE_SHEET][
            item.position_number
        ]

        logistics_block = mapping[
            LOGISTICS_SHEET
        ][item.position_number]

        economics_block = mapping[
            ECONOMICS_SHEET
        ][item.position_number]

        name_cell_value = (
            f"{item.position_number}.\n"
            f"{item.object_name}"
        )

        quantity: int | float = (
            int(item.quantity)
            if item.quantity.is_integer()
            else item.quantity
        )

        # Цены изделий:
        # A — номер/название, D — характеристики, E — количество.
        # Цены, изображения, данные поставщика и суммы не трогаем.
        set_cell(
            prices,
            price_block.first_row,
            1,
            name_cell_value,
        )
        set_cell(
            prices,
            price_block.first_row,
            5,
            quantity,
        )
        write_characteristics(
            prices,
            price_block,
            4,
            list(item.characteristics),
        )

        # Логистика:
        # A — номер/название, C — размерные характеристики,
        # D — количество, E — единица.
        set_cell(
            logistics,
            logistics_block.first_row,
            1,
            name_cell_value,
        )
        set_cell(
            logistics,
            logistics_block.first_row,
            4,
            quantity,
        )
        set_cell(
            logistics,
            logistics_block.first_row,
            5,
            item.unit,
        )
        write_characteristics(
            logistics,
            logistics_block,
            3,
            dimension_characteristics(
                item.characteristics
            ),
        )

        # Экономика: только идентификатор и количество.
        # Цены, источники товара и расчетные поля сохраняем.
        set_cell(
            economics,
            economics_block.first_row,
            1,
            name_cell_value,
        )
        set_cell(
            economics,
            economics_block.first_row,
            3,
            quantity,
        )

    # Просим Excel пересчитать сохраненные формулы при открытии.
    # Сам openpyxl значения формул не вычисляет.
    try:
        workbook.calculation = (
            openpyxl.workbook.properties.CalcProperties(
                calcMode="auto",
                fullCalcOnLoad=True,
                forceFullCalc=True,
            )
        )
    except AttributeError:
        pass


def save_workbook(
    workbook: openpyxl.Workbook,
) -> bytes:
    try:
        output = io.BytesIO()
        workbook.save(output)
        result = output.getvalue()

        if not result:
            raise ValueError(
                "Получен пустой XLSX."
            )

        return result

    except Exception as exc:
        raise AppError(
            f"Не удалось сохранить XLSX: {exc}"
        ) from exc


def items_dataframe(
    items: list[ProcurementItem],
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "№": item.position_number,
                "Объект": item.object_name,
                "Количество": item.quantity,
                "Ед. изм.": item.unit,
                "Адрес поставки": (
                    item.delivery_address
                    or "Не указан"
                ),
                "Срок поставки": (
                    item.delivery_deadline
                    or "Не указан"
                ),
                "Характеристик": len(
                    item.characteristics
                ),
            }
            for item in items
        ]
    )


# ---------------------------------------------------------------------------
# Интерфейс
# ---------------------------------------------------------------------------

def provider_settings_ui() -> LLMSettings:
    st.sidebar.header("Модель анализа")

    provider = st.sidebar.selectbox(
        "Провайдер",
        list(PROVIDERS),
    )

    config = PROVIDERS[provider]

    model_choice = st.sidebar.selectbox(
        "Модель",
        [
            *config["models"],
            "Указать модель вручную…",
        ],
        key=f"model_choice_{provider}",
    )

    if model_choice == "Указать модель вручную…":
        model = st.sidebar.text_input(
            "Идентификатор модели в API",
            key=f"custom_model_{provider}",
        ).strip()
    else:
        model = model_choice

    if provider == "Свой endpoint":
        raw_base_url = st.sidebar.text_input(
            "Base URL Chat Completions API",
            value=config["base_url"],
            help=(
                "Например, http://localhost:11434/v1. "
                "Приложение вызовет /chat/completions."
            ),
        )
    else:
        raw_base_url = config["base_url"]

        st.sidebar.caption(
            f"API endpoint: `{raw_base_url}`"
        )

    api_key = st.sidebar.text_input(
        "API-ключ",
        type="password",
        key=f"api_key_{provider}",
        help=(
            "Ключ используется только для текущих запросов. "
            "Для локального endpoint без авторизации "
            "поле можно оставить пустым."
        ),
    )

    if not model:
        raise AppError(
            "Укажите идентификатор модели."
        )

    if (
        provider != "Свой endpoint"
        and not api_key.strip()
    ):
        raise AppError(
            f"Для провайдера {provider} "
            "необходимо указать API-ключ."
        )

    return LLMSettings(
        provider=provider,
        model=model,
        base_url=normalize_base_url(
            provider,
            raw_base_url,
        ),
        api_key=api_key.strip(),
    )


def main() -> None:
    st.set_page_config(
        page_title="ТЗ → Excel",
        page_icon="📋",
        layout="wide",
    )

    st.title(
        "Техническое задание → Excel-шаблон"
    )

    st.write(
        "Выберите провайдера и модель, укажите API-ключ, "
        "загрузите PDF и подготовленный XLSX-шаблон."
    )

    try:
        settings = provider_settings_ui()
    except AppError as exc:
        st.sidebar.error(str(exc))
        return

    st.subheader("Исходные файлы")

    pdf_upload = st.file_uploader(
        "Техническое задание, PDF",
        type=["pdf"],
        key="pdf_upload",
    )

    xlsx_upload = st.file_uploader(
        "Подготовленный Excel-шаблон, XLSX",
        type=["xlsx"],
        key="xlsx_upload",
    )

    if not pdf_upload or not xlsx_upload:
        return

    try:
        pdf_bytes = read_uploaded_file(
            pdf_upload,
            MAX_PDF_BYTES,
            "PDF-файл",
        )
        xlsx_bytes = read_uploaded_file(
            xlsx_upload,
            MAX_XLSX_BYTES,
            "Excel-шаблон",
        )
    except AppError as exc:
        st.error(str(exc))
        return

    # Отпечаток зависит и от ключа: после его замены старый результат
    # не будет ошибочно показан как результат нового запуска.
    # Сам ключ не сохраняется в отпечатке в открытом виде.
    fingerprint = hashlib.sha256(
        b"\0".join(
            [
                pdf_bytes,
                xlsx_bytes,
                settings.provider.encode(),
                settings.model.encode(),
                settings.base_url.encode(),
                settings.api_key.encode(),
            ]
        )
    ).hexdigest()

    if (
        st.session_state.get(
            "result_fingerprint"
        )
        != fingerprint
    ):
        st.session_state.pop(
            "result_binary",
            None,
        )
        st.session_state.pop(
            "result_dataframe",
            None,
        )
        st.session_state.pop(
            "result_fingerprint",
            None,
        )

    already_generated = bool(
        st.session_state.get(
            "result_binary"
        )
    )

    if st.button(
        "Проанализировать и заполнить шаблон",
        type="primary",
        disabled=already_generated,
    ):
        with st.spinner(
            "Извлекаем текст, вызываем модель "
            "и проверяем шаблон…"
        ):
            try:
                pdf_text = extract_pdf_text(
                    pdf_bytes
                )

                items = extract_items_with_llm(
                    pdf_text,
                    settings,
                )

                workbook = open_template(
                    xlsx_bytes
                )

                populate_workbook(
                    workbook,
                    items,
                )

                result_binary = save_workbook(
                    workbook
                )

                st.session_state[
                    "result_binary"
                ] = result_binary

                st.session_state[
                    "result_dataframe"
                ] = items_dataframe(items)

                st.session_state[
                    "result_fingerprint"
                ] = fingerprint

            except AppError as exc:
                st.error(str(exc))

            except Exception as exc:
                st.error(
                    "Непредвиденная ошибка: "
                    f"{type(exc).__name__}: {exc}"
                )

    if st.session_state.get(
        "result_binary"
    ):
        st.success(
            "Заполненный файл подготовлен."
        )

        st.subheader(
            "Извлеченные позиции"
        )

        st.dataframe(
            st.session_state[
                "result_dataframe"
            ],
            use_container_width=True,
            hide_index=True,
        )

        st.warning(
            "Проверьте результат перед использованием. "
            "Старые цены, масса, объемы, расчеты "
            "и тексты служебных листов шаблона "
            "могут относиться к другой закупке. "
            "Формулы пересчитываются Excel "
            "при открытии файла."
        )

        st.download_button(
            label="Скачать заполненный XLSX",
            data=st.session_state[
                "result_binary"
            ],
            file_name=(
                "ТЗ_заполненный_шаблон.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


if __name__ == "__main__":
    main()