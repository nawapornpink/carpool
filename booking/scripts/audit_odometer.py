from booking.models import Car
from booking.utils.audit import month_range, audit_car_month


def run(*args):
    """
    ใช้แบบนี้:
    python manage.py runscript audit_odometer --script-args 2025 10
    """
    if len(args) < 2:
        print(
            "วิธีใช้: python manage.py runscript audit_odometer --script-args YEAR MONTH"
        )
        return

    year = int(args[0])
    month = int(args[1])

    start, end = month_range(year, month)
    print(f"🔎 Audit เดือน {month:02d}/{year} | ช่วง {start} ถึง {end}")

    for car in Car.objects.all().order_by("plate_prefix", "plate_number"):
        issues = audit_car_month(car, start, end, gap_threshold_km=200)
        if issues:
            print("\n==============================")
            print(f"🚗 รถ: {car}")
            for it in issues:
                print(" -", it["message"], "| ช่วง:", it["date_range"])
