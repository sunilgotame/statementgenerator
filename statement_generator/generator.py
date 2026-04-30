from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import date, timedelta
import random
from typing import Literal

from .utils import (
    ceil_two_decimals,
    format_amount,
    iso_date,
    next_business_day,
    parse_multiline_list,
    resolve_business_day,
    round_money,
    round_to_step,
)


EventType = Literal["deposit", "withdrawal"]

DEPOSIT_MIN_AMOUNT = 15_000
DEPOSIT_MAX_AMOUNT = 99_000
WITHDRAWAL_MIN_AMOUNT = 15_000
WITHDRAWAL_MAX_AMOUNT = 65_000

DEPOSIT_LOW_MAX_AMOUNT = 32_000
DEPOSIT_MID_MAX_AMOUNT = 65_000
DEPOSIT_HIGH_BAND_MIN_AMOUNT = 68_000
WITHDRAWAL_LOW_MAX_AMOUNT = 28_000
WITHDRAWAL_MID_MAX_AMOUNT = 42_000
WITHDRAWAL_HIGH_BAND_MIN_AMOUNT = 44_000

MIN_STATEMENT_ROWS = 40
MAX_STATEMENT_ROWS = 55

QUARTER_DEFAULT_DAYS = {
    1: 14,
    4: 13,
    7: 16,
    10: 17,
}
QUARTER_DAY_OVERRIDES = {
    (2025, 1): 13,
}


@dataclass(slots=True)
class StatementConfig:
    bank_name: str
    branch_name: str
    customer_name: str
    customer_address: str
    account_number: str
    account_type: str
    member_id: str
    currency: str
    reference_no: str
    opening_date: date | None
    start_date: date
    end_date: date
    opening_balance: float
    target_closing_balance: float
    interest_rate: float
    tax_rate: float
    cheque_start: int
    deposit_text: str
    withdrawal_text: str
    interest_text: str
    tax_text: str
    first_date_description: str = "Opening Balance"
    last_date_description: str = "Balance C/F"
    deposit_names: list[str] = field(default_factory=list)
    withdrawal_names: list[str] = field(default_factory=list)
    deposit_name_mode: str = "label_plus_name"
    withdrawal_name_mode: str = "label_plus_name"
    holiday_dates: set[date] = field(default_factory=set)
    seed: int | None = None


@dataclass(slots=True)
class PlannedEvent:
    event_type: EventType
    date: date
    amount: float


@dataclass(slots=True)
class StatementRow:
    date: date
    description: str
    cheque_no: str
    debit: float
    credit: float
    balance: float
    category: str
    is_system: bool


@dataclass(slots=True)
class StatementSummary:
    total_deposits: float = 0.0
    total_withdrawals: float = 0.0
    total_interest: float = 0.0
    total_tax: float = 0.0
    deposit_count: int = 0
    withdrawal_count: int = 0


@dataclass(slots=True)
class StatementResult:
    rows: list[StatementRow]
    summary: StatementSummary
    events: list[PlannedEvent]
    final_balance: float
    opening_business_date: date
    ending_business_date: date
    last_transaction_date: date
    issue_date: date
    seed: int


def _quarter_actual_date(year: int, month: int) -> date:
    return date(year, month, QUARTER_DAY_OVERRIDES.get((year, month), QUARTER_DEFAULT_DAYS[month]))


def _quarter_candidates(start: date, end: date) -> list[date]:
    if start > end:
        return []
    candidates: list[date] = []
    for year in range(start.year - 1, end.year + 2):
        for month in (1, 4, 7, 10):
            candidate = _quarter_actual_date(year, month)
            if start <= candidate <= end:
                candidates.append(candidate)
    return sorted(set(candidates))


def _previous_quarter_date(day_value: date) -> date | None:
    search_start = day_value - timedelta(days=400)
    candidates = [candidate for candidate in _quarter_candidates(search_start, day_value) if candidate < day_value]
    return candidates[-1] if candidates else None


def build_quarter_schedule(start: date, end: date, holidays: set[date]) -> list[tuple[date, date]]:
    del holidays
    return [(actual_date, actual_date) for actual_date in _quarter_candidates(start, end)]


def _days_in_scope(
    config: StatementConfig,
    reserved_days: set[date],
    window_start: date | None = None,
    window_end: date | None = None,
) -> dict[tuple[int, int], list[date]]:
    monthly_days: dict[tuple[int, int], list[date]] = {}
    start_bound = max(config.start_date, window_start or config.start_date)
    end_bound = min(config.end_date, window_end or config.end_date)
    current = start_bound
    while current <= end_bound:
        if current not in reserved_days and current not in config.holiday_dates:
            monthly_days.setdefault((current.year, current.month), []).append(current)
        current += timedelta(days=1)
    return monthly_days


def _allocate_weighted_counts(
    capacities: list[int],
    total_needed: int,
    rng: random.Random,
    minimums: list[int] | None = None,
) -> list[int]:
    counts = list(minimums or [0] * len(capacities))
    remaining = total_needed - sum(counts)
    if remaining < 0:
        raise ValueError("Minimum allocation exceeds total.")
    if remaining == 0:
        return counts

    base_weights = [0.85 + rng.random() * 1.15 for _ in capacities]
    guard = 0
    while remaining > 0 and guard < 50_000:
        eligible = [index for index, capacity in enumerate(capacities) if counts[index] < capacity]
        if not eligible:
            raise ValueError("Not enough capacity to allocate requested total.")
        weights: list[float] = []
        for index in eligible:
            headroom = capacities[index] - counts[index]
            weight = base_weights[index] * (0.9 + rng.random() * 0.35) * (1.0 + headroom * 0.12)
            weight /= 1.0 + (counts[index] * 0.45)
            weights.append(weight)
        chosen = rng.choices(eligible, weights=weights, k=1)[0]
        counts[chosen] += 1
        remaining -= 1
        guard += 1
    if remaining > 0:
        raise ValueError("Could not complete weighted allocation.")
    return counts


def _pick_date(
    available: list[date],
    preference: str,
    used_day_numbers: set[int],
    rng: random.Random,
) -> date:
    if not available:
        raise ValueError("No available business dates to choose from.")
    target = {"early": 0.24, "middle": 0.52, "late": 0.78}.get(preference, 0.5)
    scored: list[tuple[float, date]] = []
    length = max(len(available) - 1, 1)
    for index, day_value in enumerate(available):
        position = index / length
        score = abs(position - target)
        if day_value.day in used_day_numbers:
            score += 0.55
        score += rng.random() * 0.10
        scored.append((score, day_value))
    scored.sort(key=lambda item: item[0])
    top_choices = [day_value for _, day_value in scored[: min(5, len(scored))]]
    chosen = rng.choice(top_choices)
    available.remove(chosen)
    used_day_numbers.add(chosen.day)
    return chosen


def _estimate_net_interest(
    opening_balance: float,
    target_closing_balance: float,
    annual_rate: float,
    total_days: int,
) -> float:
    average_balance = max(0.0, (opening_balance + target_closing_balance) / 2.0)
    gross = average_balance * (annual_rate / 100.0) * (total_days / 365.0)
    return round_money(gross * 0.94)


def _random_step(rng: random.Random) -> int:
    roll = rng.random()
    if roll < 0.08:
        return 100
    if roll < 0.56:
        return 500
    return 1_000


def _random_styled_amount(minimum: int, maximum: int, rng: random.Random, forced_step: int | None = None) -> int:
    step = forced_step or _random_step(rng)
    safe_min = ((minimum + step - 1) // step) * step
    safe_max = (maximum // step) * step
    if safe_max <= safe_min:
        return safe_min
    candidates = [value for value in range(safe_min, safe_max + step, step)]
    if step == 100:
        non_five_hundred = [value for value in candidates if value % 500 != 0]
        if non_five_hundred:
            candidates = non_five_hundred
    return rng.choice(candidates)


def _distribute_total(
    total_amount: int,
    count: int,
    preferred_min: int,
    preferred_max: int,
    hard_min: int,
    hard_max: int,
    rng: random.Random,
) -> list[int]:
    if count <= 0:
        return []
    average = total_amount / count
    soft_min = max(hard_min, min(preferred_min, round_to_step(average * 0.7) or preferred_min))
    derived_soft_max = max(preferred_max, round_to_step(average * 1.35))
    soft_max = max(soft_min, min(hard_max, derived_soft_max))
    amounts: list[int] = []
    for _ in range(count):
        randomized = average * (0.75 + rng.random() * 0.53)
        bounded = max(soft_min, min(soft_max, round_to_step(randomized)))
        amounts.append(max(hard_min, min(hard_max, bounded)))

    difference = round_to_step(total_amount - sum(amounts))
    guard = 0
    while difference != 0 and guard < 20_000:
        index = guard % len(amounts)
        step = 100 if difference > 0 else -100
        next_value = amounts[index] + step
        if hard_min <= next_value <= hard_max:
            amounts[index] = next_value
            difference -= step
        guard += 1
    return [max(hard_min, min(hard_max, round_to_step(value))) for value in amounts]


def _rebalance_amounts_to_total(
    amounts: list[int],
    target_total: int,
    minimum: int,
    maximum: int,
    rng: random.Random,
    priority_order: list[int] | None = None,
) -> list[int]:
    effective_max = max(minimum, min(maximum, maximum - 1_000))
    balanced = [max(minimum, min(effective_max, round_to_step(value))) for value in amounts]
    difference = round_to_step(target_total - sum(balanced))
    guard = 0
    order = list(priority_order or list(range(len(balanced))))
    while difference != 0 and guard < 20_000:
        rng.shuffle(order)
        direction = 1 if difference > 0 else -1
        abs_difference = abs(difference)
        step_options = [step for step in (1_000, 500, 100) if step <= abs_difference]
        if not step_options:
            break

        counts = Counter(balanced)
        candidates: list[tuple[float, float, int, int]] = []
        for index in order:
            current_value = balanced[index]
            for step_size in step_options:
                next_value = current_value + (direction * step_size)
                if not (minimum <= next_value <= effective_max):
                    continue
                score = 0.0
                if step_size == 100:
                    score += 1.15
                if counts[next_value] >= 1 and next_value != current_value:
                    score += 1.5
                if counts[next_value] >= 2 and next_value != current_value:
                    score += 8.0
                if direction > 0 and next_value >= effective_max - 1_000:
                    score += 1.1
                if direction < 0 and next_value <= minimum + 1_000:
                    score += 1.1
                if counts[current_value] > 1:
                    score -= 0.5
                candidates.append((score, rng.random(), index, step_size))

        if not candidates:
            break

        candidates.sort(key=lambda item: (item[0], item[1]))
        best_score = candidates[0][0]
        top_choices = [item for item in candidates if item[0] <= best_score + 0.35][:6]
        _score, _randomizer, chosen_index, chosen_step = rng.choice(top_choices)
        balanced[chosen_index] += direction * chosen_step
        difference -= direction * chosen_step
        guard += 1
    return balanced


def _limit_duplicate_amounts(
    amounts: list[int],
    minimum: int,
    maximum: int,
    rng: random.Random,
) -> list[int]:
    if not amounts:
        return []
    effective_max = max(minimum, min(maximum, maximum - 1_000))
    adjusted = list(amounts)
    guard = 0
    while guard < 5_000:
        counts = Counter(adjusted)
        duplicate_values = [value for value, count in counts.items() if count > 1]
        overflow_value = next((value for value in duplicate_values if counts[value] > 2), None)
        if overflow_value is None and len(duplicate_values) <= 2:
            return adjusted

        if overflow_value is not None:
            candidate_index = next(index for index, value in enumerate(adjusted) if value == overflow_value and counts[value] > 2)
        else:
            duplicate_values.sort(key=lambda value: (counts[value], value), reverse=True)
            preserve = set(duplicate_values[:2])
            candidate_index = next(index for index, value in enumerate(adjusted) if value not in preserve and counts[value] > 1)

        current_value = adjusted[candidate_index]
        replacement = None
        for step_size in (1_000, 500, 100):
            offsets = [step_size, -step_size, step_size * 2, -(step_size * 2)]
            rng.shuffle(offsets)
            for offset in offsets:
                next_value = current_value + offset
                if not (minimum <= next_value <= effective_max):
                    continue
                if step_size == 100 and next_value % 500 == 0:
                    continue
                if counts[next_value] >= 2:
                    continue
                if counts[next_value] >= 1 and next_value not in duplicate_values[:2]:
                    continue
                replacement = next_value
                break
            if replacement is not None:
                break

        if replacement is None:
            guard += 1
            continue
        adjusted[candidate_index] = replacement
        guard += 1
    return adjusted


def _normalize_hundred_only_ratio(
    amounts: list[int],
    target_total: int,
    minimum: int,
    maximum: int,
    rng: random.Random,
) -> list[int]:
    if not amounts:
        return []

    effective_max = max(minimum, min(maximum, maximum - 1_000))
    adjusted = list(amounts)
    min_hundred_only = max(1, int(len(adjusted) * 0.10))
    max_hundred_only = max(min_hundred_only, int(len(adjusted) * 0.25))

    def hundred_only_indices() -> list[int]:
        return [index for index, value in enumerate(adjusted) if value % 500 != 0]

    for _ in range(4):
        hundred_indices = hundred_only_indices()
        while len(hundred_indices) > max_hundred_only:
            index = rng.choice(hundred_indices)
            current_value = adjusted[index]
            candidates = []
            for step_size in (500, 1_000):
                rounded = int(round(current_value / step_size) * step_size)
                if minimum <= rounded <= effective_max and rounded != current_value:
                    candidates.append(rounded)
            if not candidates:
                break
            adjusted[index] = min(candidates, key=lambda value: abs(value - current_value))
            hundred_indices = hundred_only_indices()

        hundred_indices = hundred_only_indices()
        while len(hundred_indices) < min_hundred_only:
            rounded_indices = [index for index, value in enumerate(adjusted) if value % 500 == 0]
            if not rounded_indices:
                break
            index = rng.choice(rounded_indices)
            current_value = adjusted[index]
            offsets = [100, 200, 300, 400, -100, -200, -300, -400]
            rng.shuffle(offsets)
            replacement = None
            for offset in offsets:
                next_value = current_value + offset
                if not (minimum <= next_value <= effective_max):
                    continue
                if next_value % 500 == 0:
                    continue
                replacement = next_value
                break
            if replacement is None:
                break
            adjusted[index] = replacement
            hundred_indices = hundred_only_indices()

        adjusted = _rebalance_amounts_to_total(adjusted, target_total, minimum, maximum, rng)
        adjusted = _limit_duplicate_amounts(adjusted, minimum, maximum, rng)

    return adjusted


def _ensure_low_band_amount(
    amounts: list[int],
    target_total: int,
    minimum: int,
    low_threshold: int,
    maximum: int,
    rng: random.Random,
) -> list[int]:
    if not amounts or any(value <= low_threshold for value in amounts):
        return amounts
    adjusted = list(amounts)
    smallest_index = min(range(len(adjusted)), key=lambda index: adjusted[index])
    low_value = _random_styled_amount(minimum, low_threshold, rng)
    adjusted[smallest_index] = low_value
    adjusted = _rebalance_amounts_to_total(adjusted, target_total, minimum, maximum, rng)
    adjusted = _limit_duplicate_amounts(adjusted, minimum, maximum, rng)
    adjusted = _normalize_hundred_only_ratio(adjusted, target_total, minimum, maximum, rng)
    return adjusted


def _apply_natural_amount_pattern(
    base_amounts: list[int],
    target_total: int,
    minimum: int,
    maximum: int,
    low_max: int,
    mid_max: int,
    high_band_min: int,
    rng: random.Random,
) -> list[int]:
    if not base_amounts:
        return []
    effective_max = max(minimum, min(maximum, maximum - 500))
    effective_high_band_min = min(effective_max - 500, high_band_min)
    styled = [max(minimum, min(effective_max, round_to_step(value))) for value in base_amounts]
    total_count = len(styled)

    low_count = min(total_count, rng.randint(min(2, total_count), min(3, total_count)))
    indices = list(range(total_count))
    rng.shuffle(indices)
    low_indices = indices[:low_count]
    remaining = [index for index in indices if index not in low_indices]
    if total_count >= 3 and remaining:
        mid_count = max(1, min(len(remaining), rng.randint(1, max(1, len(remaining)))))
    else:
        mid_count = max(0, min(len(remaining), total_count - low_count - 1))
    mid_indices = remaining[:mid_count]
    high_indices = [index for index in range(total_count) if index not in low_indices and index not in mid_indices]
    if not high_indices and total_count > 0:
        fallback_index = total_count - 1
        if fallback_index not in low_indices and fallback_index not in mid_indices:
            high_indices.append(fallback_index)

    for index in low_indices:
        styled[index] = _random_styled_amount(minimum, low_max, rng)
    for index in mid_indices:
        styled[index] = _random_styled_amount(low_max + 100, mid_max, rng)
    for index in high_indices:
        styled[index] = _random_styled_amount(max(mid_max + 100, effective_high_band_min), effective_max, rng, forced_step=500 if rng.random() < 0.5 else None)

    return _rebalance_amounts_to_total(styled, target_total, minimum, maximum, rng)


def _order_amounts_with_spacing(
    amounts: list[int],
    rng: random.Random,
    min_gap: int,
    window: int = 2,
) -> list[int]:
    remaining = list(amounts)
    ordered: list[int] = []
    while remaining:
        recent = ordered[-window:]
        scored: list[tuple[float, float, int, int]] = []
        for index, value in enumerate(remaining):
            penalty = 0.0
            for previous in recent:
                if abs(value - previous) < min_gap:
                    penalty += 1.0
                if value == previous:
                    penalty += 1.0
            scored.append((penalty, rng.random(), index, value))
        scored.sort(key=lambda item: (item[0], item[1]))
        best_penalty = scored[0][0]
        top_choices = [item for item in scored if item[0] <= best_penalty + 0.5][:4]
        chosen = rng.choice(top_choices)
        ordered.append(chosen[3])
        remaining.pop(chosen[2])
    return ordered


def _description_for_event(event: PlannedEvent, config: StatementConfig, rng: random.Random) -> str:
    if event.event_type == "deposit":
        base = config.deposit_text.strip() or "Cash Deposit"
        names = config.deposit_names
        mode = config.deposit_name_mode
    else:
        base = config.withdrawal_text.strip() or "Cheque Withdrawal"
        names = config.withdrawal_names
        mode = config.withdrawal_name_mode

    name = rng.choice(names) if names else ""
    if mode == "name_only" and name:
        return name
    if mode == "label_only" or not name:
        return base
    return f"{base} by {name}"


def _count_runs(sequence: list[EventType], event_type: EventType, length: int) -> int:
    total = 0
    current = 0
    for item in sequence:
        if item == event_type:
            current += 1
        else:
            if current == length:
                total += 1
            current = 0
    if current == length:
        total += 1
    return total


def _max_run_length(sequence: list[EventType], event_type: EventType) -> int:
    longest = 0
    current = 0
    for item in sequence:
        if item == event_type:
            current += 1
            longest = max(longest, current)
        else:
            current = 0
    return longest


def _ratio_distance(actual: float, target: float) -> float:
    return abs(actual - target)


def _resequence_transaction_types(planned: list[PlannedEvent], rng: random.Random) -> None:
    deposit_total = sum(1 for event in planned if event.event_type == "deposit")
    withdrawal_total = len(planned) - deposit_total
    total_slots = len(planned)
    if total_slots <= 1:
        return

    layouts: list[dict[str, object]] = []
    max_withdrawal_pairs = min(3, withdrawal_total // 2)
    for withdrawal_pair_runs in range(max_withdrawal_pairs + 1):
        withdrawal_runs = withdrawal_total - withdrawal_pair_runs
        if withdrawal_runs <= 0:
            continue
        withdrawal_single_runs = withdrawal_runs - withdrawal_pair_runs
        if withdrawal_single_runs < 0:
            continue

        for start_type, end_type in (
            ("deposit", "deposit"),
            ("deposit", "withdrawal"),
            ("withdrawal", "deposit"),
            ("withdrawal", "withdrawal"),
        ):
            deposit_runs = withdrawal_runs
            if start_type == "deposit":
                deposit_runs += 1
            if end_type == "withdrawal":
                deposit_runs -= 1
            if deposit_runs <= 0:
                continue

            for deposit_triple_runs in range(0, min(2, deposit_runs) + 1):
                deposit_single_runs = (2 * deposit_runs) - deposit_total + deposit_triple_runs
                deposit_double_runs = deposit_total - deposit_runs - (2 * deposit_triple_runs)
                if deposit_single_runs < 0 or deposit_double_runs < 0:
                    continue
                if deposit_single_runs + deposit_double_runs + deposit_triple_runs != deposit_runs:
                    continue
                if deposit_triple_runs > 0 and deposit_runs < 2:
                    continue
                if deposit_single_runs == 0 and deposit_triple_runs == 0 and deposit_runs > 1:
                    continue

                weight = 1.0
                if start_type == "deposit":
                    weight *= 1.15
                if end_type == "deposit":
                    weight *= 1.08
                if deposit_triple_runs == 1:
                    weight *= 1.10
                elif deposit_triple_runs == 2:
                    weight *= 0.92
                if deposit_single_runs > 0:
                    weight *= 1.10
                if withdrawal_pair_runs == 1:
                    weight *= 1.10
                elif withdrawal_pair_runs == 2:
                    weight *= 1.02
                elif withdrawal_pair_runs == 3:
                    weight *= 0.90

                layouts.append(
                    {
                        "start_type": start_type,
                        "end_type": end_type,
                        "deposit_runs": deposit_runs,
                        "withdrawal_runs": withdrawal_runs,
                        "deposit_single_runs": deposit_single_runs,
                        "deposit_double_runs": deposit_double_runs,
                        "deposit_triple_runs": deposit_triple_runs,
                        "withdrawal_pair_runs": withdrawal_pair_runs,
                        "withdrawal_single_runs": withdrawal_single_runs,
                        "weight": weight,
                    }
                )

    if not layouts:
        raise ValueError("Could not build a valid transaction run layout.")

    scored_layouts: list[dict[str, object]] = []
    for item in layouts:
        deposit_single_event_ratio = int(item["deposit_single_runs"]) / max(1, deposit_total)
        deposit_double_event_ratio = (2 * int(item["deposit_double_runs"])) / max(1, deposit_total)
        deposit_triple_event_ratio = (3 * int(item["deposit_triple_runs"])) / max(1, deposit_total)
        withdrawal_double_event_ratio = (2 * int(item["withdrawal_pair_runs"])) / max(1, withdrawal_total)

        score = 0.0
        score += _ratio_distance(deposit_single_event_ratio, 0.20) * 2.6
        score += _ratio_distance(deposit_double_event_ratio, 0.60) * 3.0
        score += _ratio_distance(deposit_triple_event_ratio, 0.20) * 2.6
        score += _ratio_distance(withdrawal_double_event_ratio, 0.30) * 2.7

        if int(item["deposit_single_runs"]) == 0:
            score += 0.7
        if int(item["deposit_triple_runs"]) == 0 and deposit_total >= 9:
            score += 0.45
        if int(item["withdrawal_pair_runs"]) == 0 and withdrawal_total >= 6:
            score += 0.25

        scored = dict(item)
        scored["score"] = score
        scored_layouts.append(scored)

    best_withdrawal_pair_score = min(
        _ratio_distance((2 * int(item["withdrawal_pair_runs"])) / max(1, withdrawal_total), 0.30)
        for item in scored_layouts
    )
    scored_layouts = [
        item
        for item in scored_layouts
        if _ratio_distance((2 * int(item["withdrawal_pair_runs"])) / max(1, withdrawal_total), 0.30)
        <= best_withdrawal_pair_score + 0.03
    ]

    best_score = min(float(item["score"]) for item in scored_layouts)
    layouts = [item for item in scored_layouts if float(item["score"]) <= best_score + 0.22]

    chosen_layout = rng.choices(
        layouts,
        weights=[float(item["weight"]) / (1.0 + (float(item["score"]) * 3.0)) for item in layouts],
        k=1,
    )[0]
    start_type = str(chosen_layout["start_type"])
    end_type = str(chosen_layout["end_type"])
    deposit_runs = int(chosen_layout["deposit_runs"])
    withdrawal_runs = int(chosen_layout["withdrawal_runs"])
    deposit_single_runs = int(chosen_layout["deposit_single_runs"])
    deposit_double_runs = int(chosen_layout["deposit_double_runs"])
    deposit_triple_runs = int(chosen_layout["deposit_triple_runs"])
    withdrawal_pair_runs = int(chosen_layout["withdrawal_pair_runs"])
    withdrawal_single_runs = int(chosen_layout["withdrawal_single_runs"])

    run_types: list[EventType] = []
    current_type: EventType = "deposit" if start_type == "deposit" else "withdrawal"
    total_runs = deposit_runs + withdrawal_runs
    for _ in range(total_runs):
        run_types.append(current_type)
        current_type = "withdrawal" if current_type == "deposit" else "deposit"
    if run_types[-1] != end_type:
        raise ValueError("Transaction run layout ended with an unexpected event type.")

    deposit_run_positions = [index for index, item in enumerate(run_types) if item == "deposit"]
    withdrawal_run_positions = [index for index, item in enumerate(run_types) if item == "withdrawal"]

    if len(deposit_run_positions) != deposit_runs or len(withdrawal_run_positions) != withdrawal_runs:
        raise ValueError("Transaction run counts do not match the selected layout.")

    run_lengths = [1] * total_runs

    if deposit_triple_runs > 0:
        candidate_positions = deposit_run_positions[:-1] if len(deposit_run_positions) > 1 else deposit_run_positions[:]
        middle_left = max(0, len(deposit_run_positions) // 4)
        middle_right = max(middle_left + 1, len(deposit_run_positions) - middle_left)
        preferred = candidate_positions[middle_left:middle_right]
        triple_pool = preferred if len(preferred) >= deposit_triple_runs else candidate_positions
        triple_positions = rng.sample(triple_pool, deposit_triple_runs)
    else:
        triple_positions = []
    for position in triple_positions:
        run_lengths[position] = 3

    remaining_deposit_positions = [position for position in deposit_run_positions if position not in triple_positions]
    single_positions: list[int] = []
    if deposit_single_runs > 0:
        if len(remaining_deposit_positions) < deposit_single_runs:
            raise ValueError("Not enough deposit run positions to place single deposit runs.")
        preferred_single_positions = [position for position in remaining_deposit_positions if position not in triple_positions]
        single_positions = rng.sample(preferred_single_positions, deposit_single_runs)
    for position in single_positions:
        run_lengths[position] = 1

    for position in remaining_deposit_positions:
        if position not in single_positions:
            run_lengths[position] = 2

    pair_positions: list[int] = []
    if withdrawal_pair_runs > 0:
        pair_positions = rng.sample(withdrawal_run_positions, withdrawal_pair_runs)
    for position in pair_positions:
        run_lengths[position] = 2
    for position in withdrawal_run_positions:
        if position not in pair_positions:
            run_lengths[position] = 1

    sequence: list[EventType] = []
    for run_type, run_length in zip(run_types, run_lengths):
        sequence.extend([run_type] * run_length)

    if len(sequence) != total_slots:
        raise ValueError("Transaction run expansion did not produce the expected number of events.")
    if sequence.count("deposit") != deposit_total or sequence.count("withdrawal") != withdrawal_total:
        raise ValueError("Transaction run expansion changed the deposit or withdrawal counts.")
    if _max_run_length(sequence, "deposit") > 3:
        raise ValueError("Transaction run expansion created a deposit streak longer than three.")
    if _max_run_length(sequence, "withdrawal") > 2:
        raise ValueError("Transaction run expansion created a withdrawal streak longer than two.")
    if _count_runs(sequence, "withdrawal", 2) > 3:
        raise ValueError("Transaction run expansion created too many double-withdrawal runs.")

    for event, new_type in zip(planned, sequence):
        event.event_type = new_type


def _plan_transaction_counts(
    monthly_days: dict[tuple[int, int], list[date]],
    system_rows: int,
    rng: random.Random,
) -> tuple[list[tuple[tuple[int, int], list[date]]], list[int], list[int], list[int], int]:
    month_items = [(month_key, days) for month_key, days in sorted(monthly_days.items()) if len(days) >= 2]
    if not month_items:
        raise ValueError("No month in the selected period has enough business days for transactions.")

    capacities = [min(4, len(days)) for _, days in month_items]
    month_count = len(month_items)
    max_user_rows = min(MAX_STATEMENT_ROWS - system_rows, sum(capacities))
    min_user_rows = min(max_user_rows, max(MIN_STATEMENT_ROWS - system_rows, month_count))
    if month_count >= 6:
        min_user_rows = min(max_user_rows, max(min_user_rows, month_count * 2))
    desired_total = max_user_rows if min_user_rows >= max_user_rows else rng.randint(min_user_rows, max_user_rows)

    monthly_totals = _allocate_weighted_counts(capacities, desired_total, rng, minimums=[1] * month_count)

    withdrawal_caps = [min(2, max(0, total - 1)) for total in monthly_totals]
    withdrawal_cap_total = sum(withdrawal_caps)
    if withdrawal_cap_total <= 0:
        raise ValueError("Could not allocate withdrawals across the selected months.")

    withdrawal_min = max(1, min(withdrawal_cap_total, round(desired_total * 0.39)))
    withdrawal_max = min(withdrawal_cap_total, max(withdrawal_min, round(desired_total * 0.45)))
    if withdrawal_max < withdrawal_min:
        withdrawal_min = min(withdrawal_cap_total, max(1, round(desired_total * 0.42)))
        withdrawal_max = withdrawal_min
    withdrawal_total = withdrawal_min if withdrawal_min >= withdrawal_max else rng.randint(withdrawal_min, withdrawal_max)
    withdrawal_counts = _allocate_weighted_counts(withdrawal_caps, withdrawal_total, rng)
    deposit_counts = [monthly_totals[index] - withdrawal_counts[index] for index in range(month_count)]
    return month_items, monthly_totals, deposit_counts, withdrawal_counts, desired_total


def _create_transaction_plan(
    config: StatementConfig,
    opening_business_date: date,
    ending_business_date: date,
    quarter_schedule: list[tuple[date, date]],
    rng: random.Random,
) -> list[PlannedEvent]:
    posting_dates = {posting_date for _, posting_date in quarter_schedule}
    reserved_days = {opening_business_date, ending_business_date, *posting_dates}
    monthly_days = _days_in_scope(config, reserved_days)
    system_rows = 2 + (len(quarter_schedule) * 2)
    month_items, _monthly_totals, deposit_counts, withdrawal_counts, _desired_total = _plan_transaction_counts(monthly_days, system_rows, rng)

    planned: list[PlannedEvent] = []
    used_day_numbers: set[int] = set()
    for index, (_month_key, available_days) in enumerate(month_items):
        working_days = list(available_days)
        for withdrawal_index in range(withdrawal_counts[index]):
            chosen = _pick_date(working_days, "late" if withdrawal_index == 0 else "middle", used_day_numbers, rng)
            planned.append(PlannedEvent("withdrawal", chosen, 0.0))
        for deposit_index in range(deposit_counts[index]):
            chosen = _pick_date(working_days, "early" if deposit_index == 0 else "middle", used_day_numbers, rng)
            planned.append(PlannedEvent("deposit", chosen, 0.0))

    planned.sort(key=lambda item: (item.date, 0 if item.event_type == "withdrawal" else 1))
    _resequence_transaction_types(planned, rng)

    deposits = [event for event in planned if event.event_type == "deposit"]
    withdrawals = [event for event in planned if event.event_type == "withdrawal"]
    total_days = max(1, (ending_business_date - opening_business_date).days + 1)
    estimated_interest = _estimate_net_interest(
        config.opening_balance,
        config.target_closing_balance,
        config.interest_rate,
        total_days,
    )

    withdrawal_amounts: list[int] = []
    if withdrawals:
        withdrawal_total_target = round_to_step(len(withdrawals) * rng.randint(22_000, 34_000))
        withdrawal_base = _distribute_total(
            withdrawal_total_target,
            len(withdrawals),
            WITHDRAWAL_MIN_AMOUNT,
            42_000,
            WITHDRAWAL_MIN_AMOUNT,
            WITHDRAWAL_MAX_AMOUNT,
            rng,
        )
        withdrawal_amounts = _apply_natural_amount_pattern(
            withdrawal_base,
            withdrawal_total_target,
            WITHDRAWAL_MIN_AMOUNT,
            WITHDRAWAL_MAX_AMOUNT,
            WITHDRAWAL_LOW_MAX_AMOUNT,
            WITHDRAWAL_MID_MAX_AMOUNT,
            WITHDRAWAL_HIGH_BAND_MIN_AMOUNT,
            rng,
        )
        withdrawal_amounts = _limit_duplicate_amounts(withdrawal_amounts, WITHDRAWAL_MIN_AMOUNT, WITHDRAWAL_MAX_AMOUNT, rng)
        withdrawal_amounts = _normalize_hundred_only_ratio(
            withdrawal_amounts,
            withdrawal_total_target,
            WITHDRAWAL_MIN_AMOUNT,
            WITHDRAWAL_MAX_AMOUNT,
            rng,
        )
        withdrawal_amounts = _order_amounts_with_spacing(withdrawal_amounts, rng, min_gap=1_200)

    deposit_amounts: list[int] = []
    if deposits:
        desired_deposit_total = max(
            round_to_step(config.target_closing_balance - config.opening_balance + sum(withdrawal_amounts) - estimated_interest),
            len(deposits) * DEPOSIT_MIN_AMOUNT,
        )
        deposit_base = _distribute_total(
            desired_deposit_total,
            len(deposits),
            32_000,
            78_000,
            DEPOSIT_MIN_AMOUNT,
            DEPOSIT_MAX_AMOUNT,
            rng,
        )
        deposit_amounts = _apply_natural_amount_pattern(
            deposit_base,
            desired_deposit_total,
            DEPOSIT_MIN_AMOUNT,
            DEPOSIT_MAX_AMOUNT,
            DEPOSIT_LOW_MAX_AMOUNT,
            DEPOSIT_MID_MAX_AMOUNT,
            DEPOSIT_HIGH_BAND_MIN_AMOUNT,
            rng,
        )
        deposit_amounts = _limit_duplicate_amounts(deposit_amounts, DEPOSIT_MIN_AMOUNT, DEPOSIT_MAX_AMOUNT, rng)
        deposit_amounts = _normalize_hundred_only_ratio(
            deposit_amounts,
            desired_deposit_total,
            DEPOSIT_MIN_AMOUNT,
            DEPOSIT_MAX_AMOUNT,
            rng,
        )
        deposit_amounts = _ensure_low_band_amount(
            deposit_amounts,
            desired_deposit_total,
            DEPOSIT_MIN_AMOUNT,
            25_000,
            DEPOSIT_MAX_AMOUNT,
            rng,
        )
        deposit_amounts = _order_amounts_with_spacing(deposit_amounts, rng, min_gap=1_500)

    for event, amount in zip(withdrawals, withdrawal_amounts):
        event.amount = float(amount)
    for event, amount in zip(deposits, deposit_amounts):
        event.amount = float(amount)

    planned.sort(key=lambda item: (item.date, 0 if item.event_type == "withdrawal" else 1))
    return planned


def _build_event_map(events: list[PlannedEvent]) -> dict[date, list[PlannedEvent]]:
    mapped: dict[date, list[PlannedEvent]] = {}
    for event in events:
        mapped.setdefault(event.date, []).append(event)
    for event_list in mapped.values():
        event_list.sort(key=lambda item: 0 if item.event_type == "withdrawal" else 1)
    return mapped


def _daily_interest(balance: float, annual_rate: float) -> float:
    return ceil_two_decimals((balance * annual_rate) / 36_500.0)


def _add_row(
    rows: list[StatementRow],
    summary: StatementSummary,
    row: StatementRow,
) -> None:
    rows.append(row)
    if row.category == "deposit":
        summary.total_deposits = round_money(summary.total_deposits + row.credit)
        summary.deposit_count += 1
    elif row.category == "withdrawal":
        summary.total_withdrawals = round_money(summary.total_withdrawals + row.debit)
        summary.withdrawal_count += 1
    elif row.category == "interest":
        summary.total_interest = round_money(summary.total_interest + row.credit)
    elif row.category == "tax":
        summary.total_tax = round_money(summary.total_tax + row.debit)


def simulate_statement(
    config: StatementConfig,
    plan: list[PlannedEvent],
    opening_business_date: date,
    ending_business_date: date,
    quarter_schedule: list[tuple[date, date]],
    rng: random.Random,
) -> tuple[list[StatementRow], StatementSummary, float, date]:
    rows: list[StatementRow] = []
    summary = StatementSummary()
    event_map = _build_event_map(plan)
    posting_dates = {posting_date for _, posting_date in quarter_schedule}

    balance = round_money(config.opening_balance)
    accrued_interest = 0.0
    previous_quarter = _previous_quarter_date(opening_business_date)
    if previous_quarter is not None:
        initial_days = min(95, max(0, (opening_business_date - previous_quarter).days))
        for _ in range(initial_days):
            accrued_interest = round_money(accrued_interest + _daily_interest(balance, config.interest_rate))

    cheque_number = config.cheque_start
    last_transaction_date = opening_business_date

    _add_row(
        rows,
        summary,
        StatementRow(
            date=opening_business_date,
            description=config.first_date_description.strip() or "Opening Balance",
            cheque_no="",
            debit=0.0,
            credit=0.0,
            balance=balance,
            category="opening",
            is_system=True,
        ),
    )

    closing_date = max([ending_business_date, *posting_dates], default=ending_business_date)
    current = opening_business_date
    while current <= closing_date:
        for event in event_map.get(current, []):
            if event.event_type == "deposit":
                balance = round_money(balance + event.amount)
                _add_row(
                    rows,
                    summary,
                    StatementRow(
                        date=current,
                        description=_description_for_event(event, config, rng),
                        cheque_no="",
                        debit=0.0,
                        credit=event.amount,
                        balance=balance,
                        category="deposit",
                        is_system=False,
                    ),
                )
            else:
                if balance - event.amount < 0:
                    raise ValueError(
                        f"Withdrawal of Rs. {format_amount(event.amount)} on {iso_date(current)} makes the balance negative."
                    )
                balance = round_money(balance - event.amount)
                _add_row(
                    rows,
                    summary,
                    StatementRow(
                        date=current,
                        description=_description_for_event(event, config, rng),
                        cheque_no=str(cheque_number),
                        debit=event.amount,
                        credit=0.0,
                        balance=balance,
                        category="withdrawal",
                        is_system=False,
                    ),
                )
                cheque_number += 1
            last_transaction_date = current

        if balance + accrued_interest > 0:
            accrued_interest = round_money(accrued_interest + _daily_interest(balance, config.interest_rate))

        if current in posting_dates:
            interest_amount = round_money(accrued_interest)
            tax_amount = round_money((interest_amount * config.tax_rate) / 100.0)
            accrued_interest = 0.0
            if interest_amount > 0:
                balance = round_money(balance + interest_amount)
                _add_row(
                    rows,
                    summary,
                    StatementRow(
                        date=current,
                        description=config.interest_text,
                        cheque_no="",
                        debit=0.0,
                        credit=interest_amount,
                        balance=balance,
                        category="interest",
                        is_system=True,
                    ),
                )
                if tax_amount > 0:
                    balance = round_money(balance - tax_amount)
                    _add_row(
                        rows,
                        summary,
                        StatementRow(
                            date=current,
                            description=config.tax_text,
                            cheque_no="",
                            debit=tax_amount,
                            credit=0.0,
                            balance=balance,
                            category="tax",
                            is_system=True,
                        ),
                    )
                last_transaction_date = current

        current += timedelta(days=1)

    _add_row(
        rows,
        summary,
        StatementRow(
            date=closing_date,
            description=config.last_date_description.strip() or "Balance C/F",
            cheque_no="",
            debit=0.0,
            credit=0.0,
            balance=balance,
            category="closing",
            is_system=True,
        ),
    )
    if len(rows) > MAX_STATEMENT_ROWS:
        raise ValueError(f"Generated statement has {len(rows)} rows, which exceeds the {MAX_STATEMENT_ROWS}-row limit.")
    return rows, summary, round_money(balance), last_transaction_date


def _latest_events_by_type(plan: list[PlannedEvent], event_type: EventType) -> list[PlannedEvent]:
    return sorted(
        [event for event in plan if event.event_type == event_type],
        key=lambda item: item.date,
        reverse=True,
    )


def _reconcile_plan(
    config: StatementConfig,
    plan: list[PlannedEvent],
    opening_business_date: date,
    ending_business_date: date,
    quarter_schedule: list[tuple[date, date]],
    rng: random.Random,
) -> tuple[list[StatementRow], StatementSummary, float, date]:
    tolerance = 100.0
    rows: list[StatementRow] = []
    summary = StatementSummary()
    final_balance = 0.0
    last_transaction_date = ending_business_date

    for _ in range(24):
        rows, summary, final_balance, last_transaction_date = simulate_statement(
            config,
            plan,
            opening_business_date,
            ending_business_date,
            quarter_schedule,
            rng,
        )
        delta = round_money(config.target_closing_balance - final_balance)
        if abs(delta) <= tolerance:
            return rows, summary, final_balance, last_transaction_date

        remaining = round_to_step(delta)
        if remaining > 0:
            deposit_events = _latest_events_by_type(plan, "deposit")
            if deposit_events:
                current_total = round_to_step(sum(event.amount for event in deposit_events))
                updated_amounts = _rebalance_amounts_to_total(
                    [int(round(event.amount)) for event in deposit_events],
                    current_total + remaining,
                    DEPOSIT_MIN_AMOUNT,
                    DEPOSIT_MAX_AMOUNT,
                    rng,
                )
                updated_amounts = _limit_duplicate_amounts(updated_amounts, DEPOSIT_MIN_AMOUNT, DEPOSIT_MAX_AMOUNT, rng)
                updated_amounts = _normalize_hundred_only_ratio(
                    updated_amounts,
                    current_total + remaining,
                    DEPOSIT_MIN_AMOUNT,
                    DEPOSIT_MAX_AMOUNT,
                    rng,
                )
                updated_amounts = _ensure_low_band_amount(
                    updated_amounts,
                    current_total + remaining,
                    DEPOSIT_MIN_AMOUNT,
                    25_000,
                    DEPOSIT_MAX_AMOUNT,
                    rng,
                )
                for event, amount in zip(deposit_events, updated_amounts):
                    event.amount = float(amount)
        else:
            needed = abs(remaining)
            deposit_events = _latest_events_by_type(plan, "deposit")
            if deposit_events:
                current_total = round_to_step(sum(event.amount for event in deposit_events))
                minimum_total = len(deposit_events) * DEPOSIT_MIN_AMOUNT
                reducible = max(0, current_total - minimum_total)
                reduction = min(needed, reducible)
                if reduction > 0:
                    updated_amounts = _rebalance_amounts_to_total(
                        [int(round(event.amount)) for event in deposit_events],
                        current_total - reduction,
                        DEPOSIT_MIN_AMOUNT,
                        DEPOSIT_MAX_AMOUNT,
                        rng,
                    )
                    updated_amounts = _limit_duplicate_amounts(updated_amounts, DEPOSIT_MIN_AMOUNT, DEPOSIT_MAX_AMOUNT, rng)
                    updated_amounts = _normalize_hundred_only_ratio(
                        updated_amounts,
                        current_total - reduction,
                        DEPOSIT_MIN_AMOUNT,
                        DEPOSIT_MAX_AMOUNT,
                        rng,
                    )
                    updated_amounts = _ensure_low_band_amount(
                        updated_amounts,
                        current_total - reduction,
                        DEPOSIT_MIN_AMOUNT,
                        25_000,
                        DEPOSIT_MAX_AMOUNT,
                        rng,
                    )
                    for event, amount in zip(deposit_events, updated_amounts):
                        event.amount = float(amount)
                    needed -= reduction
            if needed > 0:
                withdrawal_events = _latest_events_by_type(plan, "withdrawal")
                if withdrawal_events:
                    current_total = round_to_step(sum(event.amount for event in withdrawal_events))
                    updated_amounts = _rebalance_amounts_to_total(
                        [int(round(event.amount)) for event in withdrawal_events],
                        current_total + needed,
                        WITHDRAWAL_MIN_AMOUNT,
                        WITHDRAWAL_MAX_AMOUNT,
                        rng,
                    )
                    updated_amounts = _limit_duplicate_amounts(updated_amounts, WITHDRAWAL_MIN_AMOUNT, WITHDRAWAL_MAX_AMOUNT, rng)
                    updated_amounts = _normalize_hundred_only_ratio(
                        updated_amounts,
                        current_total + needed,
                        WITHDRAWAL_MIN_AMOUNT,
                        WITHDRAWAL_MAX_AMOUNT,
                        rng,
                    )
                    for event, amount in zip(withdrawal_events, updated_amounts):
                        event.amount = float(amount)

    return rows, summary, final_balance, last_transaction_date


def validate_config(config: StatementConfig) -> None:
    if not config.customer_name.strip():
        raise ValueError("Customer name is required.")
    if config.start_date >= config.end_date:
        raise ValueError("Start date must be earlier than end date.")
    if config.target_closing_balance <= 0 or config.opening_balance < 0:
        raise ValueError("Opening and closing balances must be positive values.")
    if config.target_closing_balance <= config.opening_balance:
        raise ValueError("Target closing balance must be greater than opening balance.")
    if not (0 <= config.interest_rate <= 100):
        raise ValueError("Interest rate must be between 0 and 100.")
    if not (0 <= config.tax_rate <= 100):
        raise ValueError("Tax rate must be between 0 and 100.")
    if config.cheque_start < 1:
        raise ValueError("Cheque start number must be at least 1.")


def generate_statement(config: StatementConfig) -> StatementResult:
    validate_config(config)
    seed = config.seed if config.seed is not None else random.SystemRandom().randint(10_000_000, 99_999_999)
    opening_business_date = resolve_business_day(config.start_date, config.holiday_dates)
    ending_business_date = resolve_business_day(config.end_date, config.holiday_dates)
    last_error: Exception | None = None

    for attempt in range(100):
        rng = random.Random(seed + attempt)
        try:
            quarter_schedule = build_quarter_schedule(opening_business_date, ending_business_date, config.holiday_dates)
            plan = _create_transaction_plan(config, opening_business_date, ending_business_date, quarter_schedule, rng)
            rows, summary, final_balance, last_transaction_date = _reconcile_plan(
                config,
                plan,
                opening_business_date,
                ending_business_date,
                quarter_schedule,
                rng,
            )
            if abs(final_balance - config.target_closing_balance) > 3_000:
                raise ValueError("Generated closing balance is still too far from the requested target.")
            issue_date = next_business_day(last_transaction_date, config.holiday_dates, include_self=False)
            return StatementResult(
                rows=rows,
                summary=summary,
                events=plan,
                final_balance=final_balance,
                opening_business_date=opening_business_date,
                ending_business_date=ending_business_date,
                last_transaction_date=last_transaction_date,
                issue_date=issue_date,
                seed=seed + attempt,
            )
        except Exception as error:  # pragma: no cover - guarded by retries
            last_error = error
    raise RuntimeError(str(last_error or "Unable to generate a statement for the selected inputs."))


def names_from_text(value: str) -> list[str]:
    return parse_multiline_list(value)
