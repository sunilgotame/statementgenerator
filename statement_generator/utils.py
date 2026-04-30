from __future__ import annotations

from dataclasses import asdict, is_dataclass
from datetime import date, datetime, timedelta
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable


MONTH_NAMES = (
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)

ONES = (
    "zero",
    "one",
    "two",
    "three",
    "four",
    "five",
    "six",
    "seven",
    "eight",
    "nine",
    "ten",
    "eleven",
    "twelve",
    "thirteen",
    "fourteen",
    "fifteen",
    "sixteen",
    "seventeen",
    "eighteen",
    "nineteen",
)

TENS = (
    "",
    "",
    "twenty",
    "thirty",
    "forty",
    "fifty",
    "sixty",
    "seventy",
    "eighty",
    "ninety",
)

SCALES = (
    (1_000_000_000, "billion"),
    (1_000_000, "million"),
    (1_000, "thousand"),
    (100, "hundred"),
)


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value.strip(), "%Y-%m-%d").date()


def iso_date(value: date) -> str:
    return value.strftime("%Y-%m-%d")


def format_slash_date(value: date) -> str:
    return value.strftime("%d/%m/%Y")


def ordinal_suffix(day: int) -> str:
    if 11 <= day % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def format_long_date(value: date, ordinal: bool = False) -> str:
    day_text = f"{value.day}{ordinal_suffix(value.day)}" if ordinal else str(value.day)
    return f"{day_text} {MONTH_NAMES[value.month]}, {value.year}"


def round_money(value: float) -> float:
    return round(float(value) + 1e-9, 2)


def ceil_two_decimals(value: float) -> float:
    return math.ceil((float(value) + 1e-9) * 100.0) / 100.0


def round_to_step(value: float, step: int = 100) -> int:
    return int(round(float(value) / step) * step)


def format_amount(value: float) -> str:
    return f"{round_money(value):,.2f}"


def daterange(start: date, end: date) -> Iterable[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def is_business_day(day_value: date, holidays: set[date]) -> bool:
    return day_value not in holidays


def next_business_day(day_value: date, holidays: set[date], include_self: bool = False) -> date:
    current = day_value if include_self else day_value + timedelta(days=1)
    while not is_business_day(current, holidays):
        current += timedelta(days=1)
    return current


def resolve_business_day(day_value: date, holidays: set[date]) -> date:
    if is_business_day(day_value, holidays):
        return day_value
    return next_business_day(day_value, holidays, include_self=False)


def parse_multiline_list(value: str) -> list[str]:
    items = [item.strip() for item in re.split(r"[\r\n,]+", value or "") if item.strip()]
    return items


def parse_holiday_text(value: str) -> set[date]:
    holidays: set[date] = set()
    for item in parse_multiline_list(value):
        holidays.add(parse_iso_date(item))
    return holidays


def integer_to_words(number: int) -> str:
    if number < 0:
        return f"minus {integer_to_words(abs(number))}"
    if number < 20:
        return ONES[number]
    if number < 100:
        tens, remainder = divmod(number, 10)
        return TENS[tens] if remainder == 0 else f"{TENS[tens]} {ONES[remainder]}"
    for scale_value, scale_name in SCALES:
        if number >= scale_value:
            head, tail = divmod(number, scale_value)
            if scale_value == 100:
                return (
                    f"{integer_to_words(head)} {scale_name}"
                    if tail == 0
                    else f"{integer_to_words(head)} {scale_name} {integer_to_words(tail)}"
                )
            return (
                f"{integer_to_words(head)} {scale_name}"
                if tail == 0
                else f"{integer_to_words(head)} {scale_name} {integer_to_words(tail)}"
            )
    return str(number)


def _split_money(value: float) -> tuple[int, int]:
    rounded = round_money(value)
    whole = int(math.floor(rounded))
    fraction = int(round((rounded - whole) * 100))
    if fraction == 100:
        return whole + 1, 0
    return whole, fraction


def amount_to_words_usd(value: float) -> str:
    dollars, cents = _split_money(value)
    dollar_word = "Dollar" if dollars == 1 else "Dollars"
    cent_word = "Cent" if cents == 1 else "Cents"
    return f"{integer_to_words(dollars)} {dollar_word} and {cents:02d} {cent_word} Only".title()


def amount_to_words_npr(value: float) -> str:
    rupees, paisa = _split_money(value)
    return "NPR " + f"{integer_to_words(rupees)} and {paisa:02d}/100 only".title()


def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "output"


def json_default(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, date):
        return iso_date(value)
    if is_dataclass(value):
        return asdict(value)
    raise TypeError(f"Object of type {type(value)!r} is not JSON serializable")


def write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, indent=2, default=json_default), encoding="utf-8")
