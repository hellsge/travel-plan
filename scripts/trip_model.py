"""旅行数据模型：YAML 读取、校验、写回与花销统计。"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Annotated, Any, Literal, Optional, Union, get_args

from pydantic import (
    BaseModel,
    BeforeValidator,
    ConfigDict,
    Field,
    StrictFloat,
    StrictInt,
    ValidationError,
    model_validator,
)
from ruamel.yaml import YAML
from ruamel.yaml.error import MarkedYAMLError, YAMLError
from ruamel.yaml.scalarstring import DoubleQuotedScalarString

ItemType = Literal["交通", "住宿", "游览", "餐饮", "日常", "购物", "其他"]
Status = Literal["无需预订", "待购买", "待预约", "已购买", "已预约", "已完成", "已取消"]
ITEM_TYPES: tuple[str, ...] = get_args(ItemType)
STATUSES: tuple[str, ...] = get_args(Status)
DEFAULT_STATUS = "无需预订"
TODO_STATUSES = frozenset({"待购买", "待预约"})
PAID_STATUSES = frozenset({"已购买", "已预约", "已完成"})
CANCELLED_STATUS = "已取消"
TIME_PATTERN = r"^([01]\d|2[0-3]):[0-5]\d$"
TRIP_ID_PATTERN = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
# 与站点根目录下的生成产物同名的 id 会互相覆盖。
RESERVED_IDS = frozenset({"assets", "icons"})


def _blank_to_none(value: Any) -> Any:
    if isinstance(value, str):
        value = value.strip()
        return value or None
    return value


def _strip(value: Any) -> Any:
    return value.strip() if isinstance(value, str) else value


def _parse_booking(value: Any) -> Any:
    value = _blank_to_none(value)
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(value, fmt)
            except ValueError:
                continue
            return parsed.date() if fmt == "%Y-%m-%d" else parsed
        raise ValueError("格式应为 YYYY-MM-DD 或 YYYY-MM-DD HH:MM")
    return value


OptionalText = Annotated[Optional[str], BeforeValidator(_blank_to_none)]
RequiredText = Annotated[str, BeforeValidator(_strip), Field(min_length=1)]
Clock = Annotated[
    Optional[Annotated[str, Field(pattern=TIME_PATTERN)]], BeforeValidator(_blank_to_none)
]
Money = Annotated[Optional[Union[StrictInt, StrictFloat]], Field(ge=0, allow_inf_nan=False)]
Booking = Annotated[Optional[Union[datetime, date]], BeforeValidator(_parse_booking)]


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Item(_Strict):
    title: RequiredText
    type: ItemType
    start: Clock = None
    end: Clock = None
    note: OptionalText = None
    price: Money = None
    qty: Money = None
    cost: Money = None
    booking_at: Booking = None
    status: Status = DEFAULT_STATUS
    place: OptionalText = None

    @model_validator(mode="after")
    def _cost_excludes_price(self) -> "Item":
        if self.cost is not None and (self.price is not None or self.qty is not None):
            raise ValueError("cost 不能与 price / qty 同时填写")
        return self

    @property
    def amount(self) -> Optional[float]:
        if self.cost is not None:
            return self.cost
        if self.price is not None and self.qty is not None:
            return self.price * self.qty
        return None

    @property
    def is_todo(self) -> bool:
        return self.status in TODO_STATUSES


class Day(_Strict):
    date: date
    label: OptionalText = None
    items: list[Item]


class Trip(_Strict):
    title: RequiredText
    start: date
    end: date
    travelers: Optional[int] = Field(default=None, ge=1)
    notes: OptionalText = None
    days: list[Day] = Field(min_length=1)

    @model_validator(mode="after")
    def _check_days(self) -> "Trip":
        if self.end < self.start:
            raise ValueError("end 不能早于 start")
        previous: Optional[date] = None
        for index, day in enumerate(self.days):
            if not self.start <= day.date <= self.end:
                raise ValueError(f"days[{index}].date {day.date} 超出旅行日期范围")
            if previous is not None and day.date <= previous:
                raise ValueError(f"days[{index}].date {day.date} 必须严格递增且不重复")
            previous = day.date
        return self


@dataclass
class LoadedTrip:
    id: str
    path: Path
    trip: Trip


class TripLoadError(Exception):
    def __init__(self, messages: list[str]):
        super().__init__("\n".join(messages))
        self.messages = messages


def format_loc(loc: tuple) -> str:
    text = ""
    for part in loc:
        text += f"[{part}]" if isinstance(part, int) else (f".{part}" if text else str(part))
    return text


def _describe(error: dict) -> str:
    kind, value = error["type"], error.get("input")
    if kind == "literal_error":
        return f"无效值 {value!r}（可选：{error['ctx']['expected']}）"
    if kind == "string_pattern_mismatch":
        return f"格式应为 HH:MM（加引号），实际为 {value!r}"
    if kind == "string_type":
        return f"应为字符串，实际为 {value!r}；时间请写成带引号的 \"HH:MM\""
    if kind == "extra_forbidden":
        return "未知字段"
    if kind == "missing":
        return "缺少必填字段"
    return error["msg"].removeprefix("Value error, ")


def _yaml() -> YAML:
    yaml = YAML(typ="safe", pure=True)
    yaml.version = (1, 2)
    return yaml


def load_trip(path: Path) -> LoadedTrip:
    name = path.name
    if not TRIP_ID_PATTERN.match(path.stem):
        raise TripLoadError([f"{name}: 文件名只能由小写字母、数字和 - 组成"])
    if path.stem in RESERVED_IDS:
        raise TripLoadError([f"{name}: {path.stem} 是站点保留目录名，请换一个文件名"])
    try:
        data = _yaml().load(path.read_text(encoding="utf-8"))
    except UnicodeDecodeError:
        raise TripLoadError([f"{name}: 文件不是 UTF-8 编码"]) from None
    except MarkedYAMLError as error:
        mark = error.problem_mark or error.context_mark
        line = mark.line + 1 if mark else "?"
        raise TripLoadError([f"{name}: YAML 语法错误（第 {line} 行）：{error.problem}"]) from None
    except (YAMLError, ValueError) as error:
        raise TripLoadError([f"{name}: YAML 解析失败：{error}"]) from None
    try:
        trip = Trip.model_validate(data)
    except ValidationError as error:
        messages = []
        for item in error.errors():
            loc = format_loc(item["loc"])
            messages.append(f"{name}: {loc + ': ' if loc else ''}{_describe(item)}")
        raise TripLoadError(messages) from None
    return LoadedTrip(id=path.stem, path=path, trip=trip)


def load_trips(directory: Path) -> tuple[list[LoadedTrip], list[str]]:
    trips: list[LoadedTrip] = []
    errors = [f"{p.name}: 扩展名请使用 .yaml，否则不会被发布" for p in sorted(directory.glob("*.yml"))]
    for path in sorted(directory.glob("*.yaml")):
        try:
            trips.append(load_trip(path))
        except TripLoadError as error:
            errors.extend(error.messages)
    return trips, errors


@dataclass
class TripStats:
    total: float
    per_person: Optional[float]
    by_type: list[tuple[str, float, float]]
    paid: float
    unpaid: float
    todo_count: int
    day_totals: dict[date, float]


def compute_stats(trip: Trip) -> TripStats:
    """已取消的事项不计入任何金额和待办。"""
    by_type: dict[str, float] = {}
    day_totals: dict[date, float] = {}
    paid = unpaid = 0.0
    todo_count = 0
    for day in trip.days:
        day_total = 0.0
        for item in day.items:
            if item.status == CANCELLED_STATUS:
                continue
            todo_count += item.is_todo
            amount = item.amount or 0.0
            day_total += amount
            if amount:
                by_type[item.type] = by_type.get(item.type, 0.0) + amount
            if item.status in PAID_STATUSES:
                paid += amount
            elif item.is_todo:
                unpaid += amount
        day_totals[day.date] = day_total
    total = sum(day_totals.values())
    ranked = sorted(by_type.items(), key=lambda pair: (-pair[1], ITEM_TYPES.index(pair[0])))
    return TripStats(
        total=total,
        per_person=total / trip.travelers if trip.travelers else None,
        by_type=[(kind, amount, amount / total) for kind, amount in ranked],
        paid=paid,
        unpaid=unpaid,
        todo_count=todo_count,
        day_totals=day_totals,
    )


def _plain(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    return value


def dump_trip(trip: Trip) -> str:
    """按字段定义顺序输出 YAML；时间强制加双引号，省略空值与默认状态。"""
    data = trip.model_dump(exclude_none=True)
    for day in data["days"]:
        for item in day["items"]:
            if item.get("status") == DEFAULT_STATUS:
                del item["status"]
            for key in ("start", "end"):
                if key in item:
                    item[key] = DoubleQuotedScalarString(item[key])
            for key, value in list(item.items()):
                item[key] = _plain(value)
    yaml = YAML()
    yaml.representer.ignore_aliases = lambda _data: True
    yaml.width = 4096
    yaml.indent(mapping=2, sequence=4, offset=2)
    buffer = io.StringIO()
    yaml.dump(data, buffer)
    return buffer.getvalue()
