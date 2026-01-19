# booking/utils/audit.py
from datetime import date
import calendar
from typing import List, Dict, Any, Optional

from booking.models import Booking, FuelRefill, Car


def month_range(year: int, month: int):
    """คืนค่าวันเริ่ม-วันสิ้นสุดของเดือนนั้น"""
    last = calendar.monthrange(year, month)[1]
    return date(year, month, 1), date(year, month, last)


def audit_car_month(
    car: Car, start: date, end: date, gap_threshold_km: int = 200
) -> List[Dict[str, Any]]:
    """
    ตรวจความผิดปกติของ 1 รถ ในช่วงเดือนที่กำหนด
    คืนค่า list ของ issue dict เพื่อเอาไปแสดงหน้าเว็บ/พิมพ์ใน terminal ได้
    """
    issues: List[Dict[str, Any]] = []

    # ทริปที่ “ทับช่วงเดือน”
    trips = list(
        Booking.objects.filter(
            car=car, start_date__lte=end, end_date__gte=start
        ).order_by("start_date", "id")
    )

    # เติมน้ำมันในเดือน
    fuels = list(
        FuelRefill.objects.filter(car=car, refill_date__range=(start, end)).order_by(
            "refill_date", "id"
        )
    )

    # 1) ตรวจ missing + reversed ในแต่ละทริป
    for b in trips:
        if b.odometer_before is None:
            issues.append(
                {
                    "type": "missing_before",
                    "car_id": car.id,
                    "booking_id": b.id,
                    "fuel_id": None,
                    "date_range": f"{b.start_date}..{b.end_date}",
                    "message": f"ขาดเลขไมล์ก่อนใช้งาน (ยพ. {b.yp_number})",
                }
            )

        if b.odometer_after is None:
            issues.append(
                {
                    "type": "missing_after",
                    "car_id": car.id,
                    "booking_id": b.id,
                    "fuel_id": None,
                    "date_range": f"{b.start_date}..{b.end_date}",
                    "message": f"ขาดเลขไมล์หลังใช้งาน (ยพ. {b.yp_number})",
                }
            )

        if b.odometer_before is not None and b.odometer_after is not None:
            if b.odometer_after < b.odometer_before:
                issues.append(
                    {
                        "type": "reversed_odometer",
                        "car_id": car.id,
                        "booking_id": b.id,
                        "fuel_id": None,
                        "date_range": f"{b.start_date}..{b.end_date}",
                        "message": f"เลขไมล์ถอยหลัง: before={b.odometer_before}, after={b.odometer_after} (ยพ. {b.yp_number})",
                    }
                )

    # 2) ตรวจช่องว่างระยะทาง (gap) ระหว่างทริป
    for i in range(len(trips) - 1):
        a = trips[i]
        b = trips[i + 1]
        if a.odometer_after is not None and b.odometer_before is not None:
            gap = b.odometer_before - a.odometer_after
            if gap < 0:
                issues.append(
                    {
                        "type": "gap_negative",
                        "car_id": car.id,
                        "booking_id": b.id,
                        "fuel_id": None,
                        "date_range": f"{a.end_date} -> {b.start_date}",
                        "message": f"ต่อเนื่องเลขไมล์ผิด: after ของยพ.{a.yp_number} > before ของยพ.{b.yp_number} (gap={gap})",
                    }
                )
            elif gap > gap_threshold_km:
                issues.append(
                    {
                        "type": "gap_between_trips",
                        "car_id": car.id,
                        "booking_id": b.id,
                        "fuel_id": None,
                        "date_range": f"{a.end_date} -> {b.start_date}",
                        "message": f"พบช่องว่างระยะทาง {gap} กม. ระหว่างยพ. {a.yp_number} -> {b.yp_number}",
                    }
                )

    # 3) ตรวจเติมน้ำมันกับทริปที่ผูก
    for f in fuels:
        if f.booking_id:
            bk = f.booking
            if (
                bk
                and bk.odometer_before is not None
                and f.odometer < bk.odometer_before
            ):
                issues.append(
                    {
                        "type": "fuel_odometer_outside_trip",
                        "car_id": car.id,
                        "booking_id": bk.id,
                        "fuel_id": f.id,
                        "date_range": f"{f.refill_date}",
                        "message": f"เลขไมล์ตอนเติม ({f.odometer}) < เลขไมล์ก่อนใช้งาน ({bk.odometer_before}) ของยพ.{bk.yp_number}",
                    }
                )
            if bk and bk.odometer_after is not None and f.odometer > bk.odometer_after:
                issues.append(
                    {
                        "type": "fuel_odometer_outside_trip",
                        "car_id": car.id,
                        "booking_id": bk.id,
                        "fuel_id": f.id,
                        "date_range": f"{f.refill_date}",
                        "message": f"เลขไมล์ตอนเติม ({f.odometer}) > เลขไมล์หลังใช้งาน ({bk.odometer_after}) ของยพ.{bk.yp_number}",
                    }
                )
        else:
            issues.append(
                {
                    "type": "fuel_without_booking",
                    "car_id": car.id,
                    "booking_id": None,
                    "fuel_id": f.id,
                    "date_range": f"{f.refill_date}",
                    "message": f"เติมน้ำมันวันที่ {f.refill_date} แต่ไม่ได้ผูกกับการจอง (yp_number={f.yp_number or '-'})",
                }
            )

    return issues
