from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
import json
from urllib.parse import urlencode
from urllib.request import urlopen


NRB_FOREX_API_BASE = "https://www.nrb.org.np/api/forex/v1/rates"
NRB_FOREX_PAGE = "https://www.nrb.org.np/forex/"
NRB_FOREX_DOCS = "https://www.nrb.org.np/api-docs-v1/"


class ExchangeRateLookupError(RuntimeError):
    """Raised when the exchange rate service is unavailable."""


@dataclass(slots=True)
class ExchangeRateResult:
    rate: float
    rate_type: str
    source_date: date
    source_label: str


def _request_payload(from_date: date, to_date: date, timeout: int = 20) -> dict:
    query = urlencode(
        {
            "page": 1,
            "per_page": 10,
            "from": from_date.isoformat(),
            "to": to_date.isoformat(),
        }
    )
    with urlopen(f"{NRB_FOREX_API_BASE}?{query}", timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_usd_npr_rate(issue_date: date, rate_type: str = "sell", timeout: int = 20) -> ExchangeRateResult:
    normalized_rate_type = rate_type.lower().strip()
    if normalized_rate_type not in {"buy", "sell"}:
        raise ValueError("Rate type must be either 'buy' or 'sell'.")

    for lookback in range(0, 8):
        from_date = issue_date - timedelta(days=lookback)
        payload = _request_payload(from_date, issue_date, timeout=timeout)
        entries = payload.get("data", {}).get("payload", []) or []
        for entry in reversed(entries):
            for item in entry.get("rates", []):
                currency = item.get("currency", {})
                if str(currency.get("iso3", currency.get("ISO3", ""))).upper() == "USD":
                    rate_value = float(item[normalized_rate_type])
                    source_date = date.fromisoformat(entry["date"])
                    return ExchangeRateResult(
                        rate=rate_value,
                        rate_type=normalized_rate_type,
                        source_date=source_date,
                        source_label=f"Nepal Rastra Bank {normalized_rate_type.title()} Rate",
                    )
    raise ExchangeRateLookupError(
        "Could not find a USD/NPR exchange rate from Nepal Rastra Bank for the selected issue date."
    )
