from datetime import date, timedelta
from decimal import Decimal
import random

from django.contrib.auth.models import User
from django.db import transaction
from django.db.models import Q

from booking.models import Booking, Car, FuelRefill, Profile


def run():
    with transaction.atomic():
        print("🧹 RESET ALL BOOKINGS & FUEL (YP PER CAR VERSION)")

        # ล้างข้อมูลตัวอย่าง
        FuelRefill.objects.all().delete()
        Booking.objects.all().delete()

        random.seed(12)

        # =========================
        # helper: set field if exists
        # =========================
        def set_if_exists(obj, field, value):
            if hasattr(obj, field):
                setattr(obj, field, value)

        # =========================
        # users
        # =========================
        def ensure_user(username, password="1234", first="", last="", role="EMP"):
            u, _ = User.objects.get_or_create(username=username)
            u.set_password(password)
            u.first_name = first
            u.last_name = last
            u.is_active = True
            u.save()

            p, _ = Profile.objects.get_or_create(user=u)
            if hasattr(p, "role"):
                p.role = role
            p.save()
            return u

        # ตัวหลักสำหรับ demo
        emp001 = ensure_user("emp001", "1234", "ธนดล", "แก้วพงษ์", "EMP")
        adm001 = ensure_user("adm001", "1234", "ธุรการ", "กดส", "ADM")

        # ผู้ใช้ตัวอย่างเพิ่ม
        emp_users = [
            ensure_user("emp0001", "1234", "สมศรี", "ตัวอย่าง", "EMP"),
            ensure_user("emp0002", "1234", "พรทิพย์", "ตัวอย่าง", "EMP"),
            ensure_user("emp0003", "1234", "ชยพล", "ตัวอย่าง", "EMP"),
            ensure_user("emp0004", "1234", "จิราภรณ์", "ตัวอย่าง", "EMP"),
            ensure_user("emp0005", "1234", "กิตติ", "ตัวอย่าง", "EMP"),
            ensure_user("emp0006", "1234", "สุวรรณา", "ตัวอย่าง", "EMP"),
        ]

        # requester pool: emp001 โผล่แค่บางส่วน (~25%)
        def pick_requester():
            if random.random() < 0.25:
                return emp001
            return random.choice(emp_users)

        # =========================
        # cars
        # =========================
        def get_or_create_car(prefix, number, province=None):
            car, _ = Car.objects.get_or_create(
                plate_prefix=str(prefix).strip(),
                plate_number=str(number).strip(),
            )
            if province:
                set_if_exists(car, "province_full", province)
            set_if_exists(car, "status", "READY")
            car.save()
            return car

        cars = [
            get_or_create_car("ขว", "6800", "ขอนแก่น"),
            get_or_create_car("ขว", "6808", "ขอนแก่น"),
            get_or_create_car("ขว", "6809", "ขอนแก่น"),
            get_or_create_car("งค", "3814", "ขอนแก่น"),
            get_or_create_car("งค", "3806", "ขอนแก่น"),
        ]

        # =========================
        # ✅ ตั้งเลขยพเริ่มต้นรายคัน (แก้ค่าตรงนี้ได้เลย)
        # key ใช้รูปแบบ "PREFIX-NUMBER"
        # =========================
        YP_START_BY_CAR = {
            "ขว-6800": 680001,
            "ขว-6808": 680801,
            "ขว-6809": 680901,
            "งค-3814": 381401,
            "งค-3806": 380601,
        }

        # ตัวนับเลขยพต่อคัน
        yp_counters = {}
        for c in cars:
            key = f"{c.plate_prefix}-{c.plate_number}"
            yp_counters[key] = int(YP_START_BY_CAR.get(key, 100001))

        def next_yp_for_car(car: Car) -> str:
            key = f"{car.plate_prefix}-{car.plate_number}"
            cur = int(yp_counters.get(key, 100001))
            yp_counters[key] = cur + 1
            # ให้เป็นเลข 6 หลัก (ถ้าอยากไม่ pad ก็ return str(cur))
            return str(cur).zfill(6)

        # =========================
        # overlap check
        # =========================
        def is_overlapping(car, start_d, end_d):
            return Booking.objects.filter(
                car=car,
                start_date__lte=end_d,
                end_date__gte=start_d,
            ).exists()

        # =========================
        # fuel refill
        # =========================
        PRICE_POOL = [Decimal("38.00"), Decimal("39.50"), Decimal("40.10")]
        FUEL_PLACES = ["PTT", "บางจาก", "เชลล์", "คาลเท็กซ์"]

        def add_fuel_refill(booking, odo_min, odo_max, refill_date):
            liters = Decimal(str(random.choice([15, 20, 25, 28, 30, 35])))
            ppl = random.choice(PRICE_POOL)
            total = (liters * ppl).quantize(Decimal("0.01"))
            odo = random.randint(int(odo_min), int(odo_max))

            data = dict(
                booking=booking,
                car=booking.car,
                liters=liters,
                price_per_liter=ppl,
                total_price=total,
                fuel_place=random.choice(FUEL_PLACES),
                odometer=odo,
            )
            if hasattr(FuelRefill, "refill_date"):
                data["refill_date"] = refill_date
            if hasattr(FuelRefill, "yp_number"):
                data["yp_number"] = next_yp_for_car(booking.car)

            FuelRefill.objects.create(**data)

        # =========================
        # bookings: car-centric schedule (กระจายทั้งเดือน ไม่ชน)
        # =========================
        YEAR = 2025
        MONTH = 12
        first = date(YEAR, MONTH, 1)
        last = date(YEAR, MONTH, 31)

        DESTS = [
            "กฟภ.ขอนแก่น",
            "กฟภ.อุดรธานี 2",
            "กฟภ.สกลนคร",
            "กฟภ.เลย",
            "กฟภ.หนองคาย",
            "ออกตรวจพื้นที่",
            "ประชุม/ส่งเอกสาร",
        ]

        created = 0

        for car in cars:
            # ต่อคันมี booking 7–11 รายการ
            target = random.randint(7, 11)

            cursor = first + timedelta(days=random.randint(0, 3))

            for _ in range(target):
                if cursor > last:
                    break

                length = random.choice([1, 1, 2, 2, 3])
                start_d = cursor
                end_d = min(last, start_d + timedelta(days=length - 1))

                if is_overlapping(car, start_d, end_d):
                    cursor = cursor + timedelta(days=1)
                    continue

                # สถานะสมจริง: RETURNED เยอะสุด / BOOKED บางส่วน / IN_USE นิดหน่อย / PENDING_RETURN เล็กน้อย
                r = random.random()
                if r < 0.16:
                    status = "BOOKED"
                elif r < 0.30:
                    status = "IN_USE"
                elif r < 0.38:
                    status = "PENDING_RETURN"  # ✅ เคสที่รอเติมน้ำมัน/รอยืนยัน
                else:
                    status = "RETURNED"

                requester = pick_requester()

                b = Booking.objects.create(
                    car=car,
                    requester=requester,
                    start_date=start_d,
                    end_date=end_d,
                    destination=random.choice(DESTS),
                    status=status,
                    returned_by=(
                        adm001
                        if status == "RETURNED"
                        else (requester if status == "PENDING_RETURN" else None)
                    ),
                )

                # เลขไมล์
                if status in ("IN_USE", "RETURNED", "PENDING_RETURN"):
                    odo_before = random.randint(20000, 90000)
                    odo_after = odo_before + random.randint(40, 450)

                    set_if_exists(b, "odometer_before", odo_before)
                    set_if_exists(b, "mileage_before", odo_before)

                    if status in ("RETURNED", "PENDING_RETURN"):
                        set_if_exists(b, "odometer_after", odo_after)
                        set_if_exists(b, "mileage_after", odo_after)

                    b.save()

                    # เติมน้ำมัน: 0–3 ครั้งต่อเคส (ให้เหมือนไปราชการจริง)
                    # RETURNED: เติมบ่อยหน่อย
                    # PENDING_RETURN: มีโอกาสเติมแล้วแต่ยังไม่กดยืนยัน
                    if status in ("RETURNED", "PENDING_RETURN"):
                        n = random.choices([0, 1, 2, 3], weights=[20, 45, 25, 10])[0]
                    else:
                        n = random.choices([0, 1, 2], weights=[55, 35, 10])[0]

                    day_span = max(0, (end_d - start_d).days)
                    for _i in range(n):
                        refill_date = start_d + timedelta(
                            days=random.randint(0, day_span)
                        )
                        add_fuel_refill(b, odo_before, odo_after, refill_date)

                created += 1

                gap = random.choice([0, 1, 1, 2])
                cursor = end_d + timedelta(days=1 + gap)

        # เซ็ตสถานะรถให้ READY ไว้ก่อน (รถพร้อมโชว์ในหน้าเลือก)
        for c in cars:
            set_if_exists(c, "status", "READY")
            c.save()

        print("✅ DONE")
        print("   - bookings:", created)
        print("   - fuel refills:", FuelRefill.objects.count())
        print("   - yp starts:", YP_START_BY_CAR)
