from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo


_STT_PATTERNS = (
    re.compile(r"^\s*\[(\d{1,3})\](?:\s|[-._:|]|$)"),
    re.compile(r"^\s*\((\d{1,3})\)(?:\s|[-._:|]|$)"),
    re.compile(r"^\s*(\d{1,3})(?=\s|[-._:|]|$)"),
)


@dataclass(frozen=True)
class Shift:
    key: str
    label: str
    start: time
    end: time
    minimum_minutes: float
    excel_offset: int


def parse_hhmm(raw: str) -> time:
    """Đổi HH:MM thành datetime.time và báo lỗi rõ ràng nếu nhập sai."""
    value = raw.strip()
    pieces = value.split(":")
    if len(pieces) != 2 or not all(piece.isdigit() for piece in pieces):
        raise ValueError(f"Giờ không hợp lệ: {raw!r}. Hãy dùng HH:MM, ví dụ 20:00.")
    hour, minute = map(int, pieces)
    if not 0 <= hour <= 23 or not 0 <= minute <= 59:
        raise ValueError(f"Giờ không hợp lệ: {raw!r}.")
    return time(hour, minute)


def extract_stt(display_name: str) -> int | None:
    """Đọc STT ở đầu tên: 12 Tên, 12-Tên, [12] Tên hoặc (12) Tên."""
    for pattern in _STT_PATTERNS:
        match = pattern.search(display_name or "")
        if match:
            number = int(match.group(1))
            return number if number > 0 else None
    return None


def shift_window(shift_day: date, shift: Shift, tz: ZoneInfo) -> tuple[datetime, datetime]:
    """Tạo khoảng giờ ca cố định, kể cả khi ca đi qua nửa đêm."""
    start = datetime.combine(shift_day, shift.start, tzinfo=tz)
    end = datetime.combine(shift_day, shift.end, tzinfo=tz)
    if end <= start:
        end += timedelta(days=1)
    return start, end


def attendance_column(weekday: int, excel_offset: int, start_column: int = 5) -> int:
    """E=5 là Boss Trưa Thứ 2; mỗi ngày có 3 cột."""
    if weekday not in range(7):
        raise ValueError("weekday phải từ 0 (Thứ 2) đến 6 (Chủ nhật)")
    if excel_offset not in range(3):
        raise ValueError("excel_offset phải là 0, 1 hoặc 2")
    return start_column + weekday * 3 + excel_offset


def manual_required_seconds(duration_seconds: float, ratio: float = 3 / 5) -> float:
    """Ca linh động yêu cầu có mặt đúng 3/5 tổng thời lượng ca."""
    if duration_seconds < 0:
        raise ValueError("Thời lượng ca không được âm")
    if not 0 < ratio <= 1:
        raise ValueError("Tỷ lệ điểm danh phải nằm trong khoảng (0, 1]")
    return duration_seconds * ratio


def format_duration(seconds: float) -> str:
    seconds = max(0, round(seconds))
    minutes, second = divmod(seconds, 60)
    hours, minute = divmod(minutes, 60)
    if hours:
        return f"{hours} giờ {minute:02d} phút {second:02d} giây"
    return f"{minute} phút {second:02d} giây"
