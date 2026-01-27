# ============================================================
# booking/views.py  (refactored: organized, deduped imports)
# ============================================================

from __future__ import annotations

# -----------------------------
# Standard library
# -----------------------------
import os
import calendar
import zipfile
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO

# -----------------------------
# Django
# -----------------------------
from django.conf import settings
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required, user_passes_test
from django.contrib.auth.models import User
from django.db.models import Q
from django.http import HttpResponse, HttpResponseForbidden, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.dateparse import parse_date
from django.views.decorators.http import require_POST
from django.db.models import Case, When, Value, IntegerField

import booking
import json

# -----------------------------
# Local
# -----------------------------
from .forms import EmployeeCreateForm, EmployeeUpdateForm, CarForm, FuelRefillForm
from .models import Booking, Car, FuelRefill, Profile
from .services.report_fuel_excel import build_fuel_excel
from .services.report_car_docx import build_car_docx
from urllib.parse import quote

# =============================================================================
# Helpers
# =============================================================================


def get_profile(user: User) -> Profile | None:
    try:
        return user.profile
    except Exception:
        return None


def is_admin(user: User) -> bool:
    p = get_profile(user)
    return bool(p and p.role == "ADM") or user.is_staff or user.is_superuser


def admin_required(view_func):
    def _wrapped(request, *args, **kwargs):
        # ยังไม่ได้ login
        if not request.user.is_authenticated:
            # ถ้าเรียกมาจาก fetch/ajax ให้คืน JSON
            if request.headers.get("X-Requested-With") == "XMLHttpRequest":
                return JsonResponse({"error": "not_authenticated"}, status=401)
            return redirect("booking:login")

        # login แล้ว แต่ไม่ใช่ admin
        if not is_admin(request.user):
            if request.headers.get("X-Requested-With") == "XMLHttpRequest":
                return JsonResponse({"error": "forbidden"}, status=403)
            return HttpResponseForbidden("Forbidden")

        return view_func(request, *args, **kwargs)

    return _wrapped


def is_booking_member(booking: Booking, user: User) -> bool:
    """ผู้ที่ทำรายการคืนรถ/ดูรายการได้: ผู้สร้างคำขอ หรืออยู่ในผู้เดินทาง"""
    if booking.requester_id == user.id:
        return True
    try:
        prof = user.profile
    except Exception:
        return False
    return booking.co_travelers.filter(id=prof.id).exists()


def safe_decimal(val: str | None) -> Decimal | None:
    if val is None:
        return None
    val = str(val).strip()
    if val == "":
        return None
    try:
        return Decimal(val)
    except InvalidOperation:
        return None


def _to_decimal(val):
    val = (val or "").strip()
    if val == "":
        return None
    try:
        return Decimal(val)
    except (InvalidOperation, ValueError):
        return None


# =============================================================================
# Auth
# =============================================================================


def login_view(request):
    if request.method == "POST":
        employee_id = (request.POST.get("employee_id") or "").strip()
        password = (request.POST.get("password") or "").strip()

        user = authenticate(request, username=employee_id, password=password)
        if user is not None:
            login(request, user)
            return redirect("booking:dashboard_redirect")

        messages.error(request, "รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง")
    return render(request, "booking/login.html")


@login_required
def logout_view(request):
    logout(request)
    return redirect("booking:login")


@login_required
def dashboard_redirect(request):
    if is_admin(request.user):
        return redirect("booking:admin_dashboard")
    return redirect("booking:user_dashboard")


# =============================================================================
# USER: Dashboard + Calendar
# =============================================================================


@login_required
def user_dashboard(request):
    profile = get_profile(request.user)
    today = timezone.localdate()

    # รถพร้อมให้จอง = status READY และ "ไม่ได้ถูกจอง/ใช้งานทับวันนี้"
    busy_car_ids_today = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .filter(start_date__lte=today, end_date__gte=today)
        .values_list("car_id", flat=True)
        .distinct()
    )
    available_cars = (
        Car.objects.filter(status="READY")
        .exclude(id__in=list(busy_car_ids_today))
        .order_by("plate_prefix", "plate_number")
    )

    # ปฏิทิน/dropdown ต้องใช้รถทั้งหมด
    all_cars = Car.objects.all().order_by("plate_prefix", "plate_number")

    # ปฏิทินในหน้า dashboard ใช้ all_bookings (ไม่เอา CANCELLED)
    # ปฏิทินในหน้า dashboard: เอาเฉพาะที่ยังจอง/กำลังใช้งาน
    all_bookings = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .select_related("car", "requester")
        .prefetch_related("co_travelers", "co_travelers__user")
        .order_by("start_date")
    )

    # ตาราง "รายการจองของคุณ"
    my_bookings = (
        Booking.objects.filter(
            Q(requester=request.user) | Q(co_travelers__user=request.user)
        )
        .distinct()
        .select_related("car", "requester")
        .order_by("-start_date", "-created_at")
    )

    # ✅ สิทธิ์ยกเลิก: ยกเลิกได้ถึงนาทีสุดท้าย
    # เงื่อนไขเดียว: ยังเป็น BOOKED
    for b in my_bookings:
        b.can_cancel = (b.status == "BOOKED")

    return render(
        request,
        "booking/user_dashboard.html",
        {
            "profile": profile,
            "available_cars": available_cars,
            "all_cars": all_cars,
            "all_bookings": all_bookings,
            "my_bookings": my_bookings,
        },
    )


@login_required
def user_calendar(request):
    """
    หน้า calendar แบบ grid (user_calendar.html) ที่ใช้ weeks + bookings_by_day
    """
    today = timezone.localdate()
    year = int(request.GET.get("year") or today.year)
    month = int(request.GET.get("month") or today.month)
    car_id = request.GET.get("car_id") or ""

    cal = calendar.Calendar(firstweekday=0)
    month_days = cal.monthdatescalendar(year, month)

    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    qs = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .filter(start_date__lte=last_day, end_date__gte=first_day)
        .select_related("car")
        .order_by("start_date")
    )

    if car_id:
        qs = qs.filter(car_id=car_id)

    bookings_by_day = {d: [] for week in month_days for d in week}

    for b in qs:
        cur = b.start_date
        while cur <= b.end_date:
            if cur in bookings_by_day:
                bookings_by_day[cur].append(b)
            cur += timedelta(days=1)

    cars = Car.objects.all().order_by("plate_prefix", "plate_number")

    return render(
        request,
        "booking/user_calendar.html",
        {
            "weeks": month_days,
            "year": year,
            "month": month,
            "cars": cars,
            "selected_car_id": car_id,
            "bookings_by_day": bookings_by_day,
        },
    )


@login_required
def user_booking_events(request):
    """
    API สำหรับ FullCalendar (ถ้าอยากย้ายไปโหลดผ่าน AJAX)
    """
    qs = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .select_related("car")
        .order_by("start_date")
    )

    events = []
    for b in qs:
        if not b.car_id:
            continue
        events.append(
            {
                "id": b.id,
                "title": f"{b.car.plate_prefix}{b.car.plate_number} {b.car.province_full}",
                "start": b.start_date.isoformat(),
                "end": (b.end_date + timedelta(days=1)).isoformat(),
                "allDay": True,
                "backgroundColor": b.car.color_code or "#6366f1",
                "borderColor": b.car.color_code or "#6366f1",
                "extendedProps": {
                    "carId": b.car_id,
                    "destination": b.destination or "",
                    "status": b.status,
                },
            }
        )

    return JsonResponse(events, safe=False)


# =============================================================================
# USER: Booking create
# =============================================================================


@login_required
def api_available_cars(request):
    """
    API รถที่พร้อมให้จอง (ใช้ในฟอร์มจอง)
    รับ start_date, end_date แล้วคืนรถ READY ที่ไม่มี booking ทับช่วงนั้น
    """
    start = parse_date(request.GET.get("start_date") or "")
    end = parse_date(request.GET.get("end_date") or "")

    if not start or not end:
        return JsonResponse({"ok": False, "error": "missing_dates"}, status=400)
    if end < start:
        return JsonResponse({"ok": False, "error": "end_before_start"}, status=400)

    busy_car_ids = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .filter(start_date__lte=end, end_date__gte=start)
        .values_list("car_id", flat=True)
        .distinct()
    )

    cars = (
        Car.objects.filter(status="READY")
        .exclude(id__in=list(busy_car_ids))
        .order_by("plate_prefix", "plate_number")
    )

    data = [
        {
            "id": c.id,
            "plate_prefix": c.plate_prefix,
            "plate_number": c.plate_number,
            "province_full": c.province_full,
            "brand_name": c.brand_name,
            "model_name": c.model_name,
            "color_code": c.color_code,
        }
        for c in cars
    ]
    return JsonResponse({"ok": True, "cars": data})


@login_required
def user_create_booking(request):
    """
    ฟอร์มจองรถ: user_booking_form.html
    """
    cars = Car.objects.filter(status="READY").order_by("plate_prefix", "plate_number")
    employees = Profile.objects.select_related("user").order_by(
        "user__first_name", "user__last_name", "user__username"
    )

    if request.method == "POST":
        car_id = request.POST.get("car") or ""
        start = parse_date(request.POST.get("start_date") or "")
        end = parse_date(request.POST.get("end_date") or "")
        destination = (request.POST.get("destination") or "").strip()
        co_travelers_ids = request.POST.getlist("co_travelers")

        if not car_id or not start or not end or not destination:
            messages.error(request, "กรอกข้อมูลให้ครบ: รถ, วันไป-วันกลับ, ปลายทาง")
            return render(
                request,
                "booking/user_booking_form.html",
                {"cars": cars, "employees": employees},
            )

        if end < start:
            messages.error(request, "วันกลับต้องไม่ก่อนวันไป")
            return render(
                request,
                "booking/user_booking_form.html",
                {"cars": cars, "employees": employees},
            )

        car = get_object_or_404(Car, id=car_id)

        conflict = (
            Booking.objects.filter(car=car, status__in=["BOOKED", "IN_USE"])
            .filter(start_date__lte=end, end_date__gte=start)
            .exists()
        )
        if conflict:
            messages.error(request, "รถคันนี้มีการจองทับช่วงวันดังกล่าวแล้ว กรุณาเลือกวันหรือรถใหม่")
            return render(
                request,
                "booking/user_booking_form.html",
                {"cars": cars, "employees": employees},
            )
        booking = Booking.objects.create(
            car=car,
            requester=request.user,
            start_date=start,
            end_date=end,
            destination=destination,
            status="BOOKED",
        )

        if co_travelers_ids:
            booking.co_travelers.set(Profile.objects.filter(id__in=co_travelers_ids))

        messages.success(request, "จองรถเรียบร้อย")
        return redirect("booking:user_dashboard")

    return render(
        request,
        "booking/user_booking_form.html",
        {"cars": cars, "employees": employees},
    )


@login_required
def ajax_available_cars(request):
    start_date = parse_date(request.GET.get("start_date"))
    end_date = parse_date(request.GET.get("end_date"))

    if not start_date or not end_date:
        return JsonResponse({"ok": False, "cars": []})

    # เงื่อนไขชนช่วงวัน (overlap):
    # มี booking ที่ start <= end_date และ end >= start_date
    overlapping = Booking.objects.filter(
        status__in=["BOOKED", "IN_USE"],  # ปรับตามสถานะที่ “กันรถ”
        start_date__lte=end_date,
        end_date__gte=start_date,
    ).values_list("car_id", flat=True)

    cars = (
        Car.objects.filter(status="READY")
        .exclude(id__in=overlapping)
        .order_by("plate_prefix", "plate_number")
    )

    return JsonResponse(
        {
            "ok": True,
            "cars": [
                {
                    "id": c.id,
                    "label": f"{c.plate_prefix} {c.plate_number} ({c.brand_name} {c.model_name})",
                }
                for c in cars
            ],
        }
    )


# =============================================================================
# USER: Cancel booking
# =============================================================================


@login_required
@require_POST
def user_cancel_booking(request, booking_id):
    """
    ยกเลิกการจอง (เฉพาะ POST)
    - ผู้สร้างคำขอ หรืออยู่ในผู้เดินทางเท่านั้น
    - ยกเลิกได้เฉพาะสถานะ BOOKED
    - ✅ ยกเลิกได้ถึงนาทีสุดท้าย (ไม่จำกัดก่อนวันเริ่ม 1 วันแล้ว)
    """

    booking = get_object_or_404(
        Booking.objects
            .select_related("car", "requester")
            .prefetch_related("fuel_refills"),
        id=booking_id,
    )

    # ✅ ต้องเป็นผู้สร้างคำขอหรืออยู่ในผู้เดินทาง
    if not is_booking_member(booking, request.user):
        return HttpResponseForbidden("Forbidden")

    # ✅ ยกเลิกได้เฉพาะ BOOKED
    if booking.status != "BOOKED":
        messages.warning(
            request,
            "รายการนี้ไม่สามารถยกเลิกได้ (อาจเริ่มใช้งานแล้ว / คืนแล้ว / ถูกยกเลิกไปแล้ว)"
        )
        return redirect("booking:user_dashboard")

    booking.status = "CANCELLED"
    booking.save(update_fields=["status", "updated_at"])

    messages.success(
        request,
        f"ยกเลิกการจองรถ {booking.car.plate_prefix} {booking.car.plate_number} เรียบร้อยแล้ว"
    )
    return redirect("booking:user_dashboard")

# =============================================================================
# USER: Return car flow (ถามเติมน้ำมันก่อน)
# =============================================================================


@login_required
def user_return_car(request):
    """
    หน้าเริ่มใช้งาน/คืนรถ: user_return_car.html
    - start_use: กรอก odometer_before -> status IN_USE
    - return:
        ส่ง has_fuel = YES/NO
        - NO: คืนรถทันที (status RETURNED)
        - YES: เก็บ odometer_after ไว้ใน session แล้ว redirect ไปหน้าเติมน้ำมัน (booking_id)
    """
    available_bookings = (
        Booking.objects.exclude(status__in=["RETURNED", "CANCELLED"])
        .filter(Q(requester=request.user) | Q(co_travelers__user=request.user))
        .distinct()
        .select_related("car", "requester")
        .order_by("-start_date", "-created_at")
    )

    booking = None
    selected_booking_id = None
    booking_id = request.GET.get("booking_id")
    if booking_id:
        booking = get_object_or_404(
            Booking.objects.select_related("car", "requester"),
            id=booking_id
        )
        selected_booking_id = booking.id

    if request.method == "POST":
        action = request.POST.get("action") or ""

        if not booking:
            messages.error(request, "กรุณาเลือกรายการก่อน")
            return render(
                request,
                "booking/user_return_car.html",
                {"available_bookings": available_bookings, "booking": booking},
            )

        if action == "start_use":
            if booking.status != "BOOKED":
                messages.error(request, "รายการนี้ไม่อยู่ในสถานะที่เริ่มใช้งานได้")
                return redirect(f"{request.path}?booking_id={booking.id}")

            odo_before = request.POST.get("odometer_before")
            try:
                odo_before_int = int(odo_before)
            except (TypeError, ValueError):
                messages.error(request, "เลขไมล์ก่อนต้องเป็นตัวเลข")
                return redirect(f"{request.path}?booking_id={booking.id}")

            booking.odometer_before = odo_before_int
            booking.status = "IN_USE"
            booking.save(update_fields=["odometer_before", "status", "updated_at"])

            if booking.car:
                booking.car.current_odometer = max(
                    booking.car.current_odometer or 0, odo_before_int
                )
                booking.car.save(update_fields=["current_odometer"])

            messages.success(request, "เริ่มใช้งานเรียบร้อย")
            return redirect(f"{request.path}?booking_id={booking.id}")

        if action == "return":
            if booking.status != "IN_USE":
                messages.error(request, "รายการนี้ไม่อยู่ในสถานะที่คืนรถได้")
                return redirect(f"{request.path}?booking_id={booking.id}")

            odo_after = request.POST.get("odometer_after")
            try:
                odo_after_int = int(odo_after)
            except (TypeError, ValueError):
                messages.error(request, "เลขไมล์หลังต้องเป็นตัวเลข")
                return redirect(f"{request.path}?booking_id={booking.id}")

            if (
                booking.odometer_before is not None
                and odo_after_int < booking.odometer_before
            ):
                messages.error(request, "เลขไมล์หลังต้องไม่ต่ำกว่าเลขไมล์ก่อน")
                return redirect(f"{request.path}?booking_id={booking.id}")

            has_fuel = (request.POST.get("has_fuel") or "").strip().upper()
            if has_fuel not in ["YES", "NO"]:
                messages.error(request, "กรุณาเลือกว่ามีการเติมน้ำมันหรือไม่ก่อนคืนรถ")
                return redirect(f"{request.path}?booking_id={booking.id}")

            if has_fuel == "NO":
                booking.odometer_after = odo_after_int
                booking.status = "RETURNED"
                booking.returned_by = request.user
                booking.save(
                    update_fields=[
                        "odometer_after",
                        "status",
                        "returned_by",
                        "updated_at",
                    ]
                )

                if booking.car:
                    booking.car.current_odometer = max(
                        booking.car.current_odometer or 0, odo_after_int
                    )
                    booking.car.save(update_fields=["current_odometer"])

                messages.success(request, "คืนรถเรียบร้อยแล้ว (ไม่มีเติมน้ำมัน)")
                return redirect("booking:user_return_car")

            # YES -> ตั้งสถานะรอคืนรถ (ยังไม่ RETURNED)
            booking.odometer_after = odo_after_int
            booking.status = "PENDING_RETURN"
            booking.returned_by = request.user
            booking.save(
                update_fields=["odometer_after", "status", "returned_by", "updated_at"]
            )

            return redirect(
                f"{redirect('booking:user_fuel_refill').url}?booking_id={booking.id}"
            )

    return render(
        request,
        "booking/user_return_car.html",
        {
            "available_bookings": available_bookings,
            "booking": booking,
            "selected_booking_id": selected_booking_id,  # ✅ เพิ่ม
        },
    )



# =============================================================================
# USER: Fuel refill (รองรับมาจาก return flow)
# =============================================================================


@login_required
def user_fuel_refill(request):
    # --- 1) หา booking (ถ้ามาจากคืนรถจะมี booking_id) ---
    booking = None
    booking_id = request.GET.get("booking_id") or request.POST.get("booking_id") or ""
    if booking_id:
        try:
            b = Booking.objects.select_related("car", "requester").get(id=booking_id)
            if is_booking_member(b, request.user):
                booking = b
        except Booking.DoesNotExist:
            booking = None

    cars = Car.objects.all().order_by("plate_prefix", "plate_number")

    # ✅ ตารางล่าง: โชว์เฉพาะเคสที่เลือกเท่านั้น (กันมั่ว)
    refills_qs = FuelRefill.objects.none()
    if booking:
        refills_qs = (
            FuelRefill.objects.filter(booking=booking)
            .select_related("car")
            .order_by("-created_at")
        )

    # --- 2) GET: แสดงฟอร์ม ---
    if request.method != "POST":
        return render(
            request,
            "booking/user_fuel_refill.html",
            {
                "cars": cars,
                "booking": booking,
                "refills": refills_qs,
                "today": timezone.localdate(),
            },
        )

    # --- 3) POST: บันทึกเติมน้ำมัน ---
    car_id = request.POST.get("car_id")
    yp_number = (request.POST.get("yp_number") or "").strip()
    fuel_place = (request.POST.get("fuel_place") or "").strip()

    refill_date = parse_date(request.POST.get("refill_date") or "")
    if not refill_date:
        refill_date = timezone.localdate()

    odometer = request.POST.get("odometer")
    price_per_liter = _to_decimal(request.POST.get("price_per_liter"))
    liters = _to_decimal(request.POST.get("liters"))
    total_price = _to_decimal(request.POST.get("total_price"))

    # ✅ helper redirect กลับหน้าเดิมให้ไม่หลุด booking_id
    def back():
        if booking:
            return redirect(
                f"{redirect('booking:user_fuel_refill').url}?booking_id={booking.id}"
            )
        return redirect("booking:user_fuel_refill")

    if not car_id:
        messages.error(request, "กรุณาเลือกทะเบียนรถ")
        return back()

    if not fuel_place:
        messages.error(request, "กรุณากรอกสถานที่เติมน้ำมัน 697")
        return back()

    if not yp_number:
        messages.error(request, "กรุณากรอกเลข ยพ.")
        return back()

    if not odometer:
        messages.error(request, "กรุณากรอกเลขไมล์")
        return back()

    try:
        odometer_int = int(odometer)
    except (TypeError, ValueError):
        messages.error(request, "เลขไมล์ต้องเป็นตัวเลข 711")
        return back()

    if price_per_liter is None:
        messages.error(request, "กรุณากรอกราคาน้ำมันต่อลิตรให้ถูกต้อง")
        return back()

    if liters is None:
        messages.error(request, "กรุณากรอกจำนวนลิตรให้ถูกต้อง")
        return back()

    if total_price is None:
        messages.error(request, "กรุณากรอกราคารวมให้ถูกต้อง")
        return back()

    car = get_object_or_404(Car, id=car_id)

    FuelRefill.objects.create(
        booking=booking,  # ต้องมี booking_id ถึงจะเป็นเคสต่อเคส
        car=car,
        refill_date=refill_date,
        fuel_place=fuel_place,
        yp_number=yp_number,
        odometer=odometer_int,
        price_per_liter=price_per_liter,
        liters=liters,
        total_price=total_price,
    )

    # ✅ ไม่คืนรถอัตโนมัติแล้ว (ลบ session auto-return ทิ้ง)
    messages.success(request, "บันทึกการเติมน้ำมันเรียบร้อย")
    return redirect(
        f"{redirect('booking:user_fuel_refill').url}?booking_id={booking.id}"
    )


@login_required
def user_booking_detail(request, booking_id: int):
    booking = get_object_or_404(
        Booking.objects.select_related(
            "car", "requester", "returned_by"
        ).prefetch_related("co_travelers", "co_travelers__user"),
        id=booking_id,
    )

    # ✅ ให้ผู้สร้างคำขอ + ผู้เดินทางดูได้
    if not is_booking_member(booking, request.user):
        return HttpResponseForbidden("Forbidden")

    fuels = (
        FuelRefill.objects.filter(booking=booking)
        .select_related("car")
        .order_by("-created_at")
    )

    return render(
        request,
        "booking/user_booking_detail.html",
        {"booking": booking, "fuels": fuels},
    )


@login_required
@require_POST
def user_confirm_return(request, booking_id):
    booking = get_object_or_404(Booking, id=booking_id)

    if not is_booking_member(booking, request.user):
        return HttpResponseForbidden("Forbidden")

    if booking.status != "PENDING_RETURN":
        messages.error(request, "สถานะไม่ถูกต้อง")
        return redirect(
            f"{redirect('booking:user_fuel_refill').url}?booking_id={booking.id}"
        )

    if not FuelRefill.objects.filter(booking=booking).exists():
        messages.error(request, "ต้องมีรายการเติมน้ำมันอย่างน้อย 1 รายการ")
        return redirect(
            f"{redirect('booking:user_fuel_refill').url}?booking_id={booking.id}"
        )

    booking.status = "RETURNED"
    booking.save(update_fields=["status", "updated_at"])

    if booking.car and booking.odometer_after is not None:
        booking.car.current_odometer = max(
            booking.car.current_odometer or 0, booking.odometer_after
        )
        booking.car.status = "READY"
        booking.car.save(update_fields=["current_odometer", "status"])

    messages.success(request, "ยืนยันคืนรถเรียบร้อย")
    return redirect("booking:user_dashboard")


# =============================================================================
# ADMIN: Dashboard / Employees / Cars
# =============================================================================


@admin_required
def admin_dashboard(request):
    today = timezone.localdate()

    total_cars = Car.objects.count()
    total_users = User.objects.count()

    active_bookings = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .filter(start_date__lte=today, end_date__gte=today)  # ✅ เฉพาะของวันนี้
        .select_related("car", "requester")
        .order_by("car__plate_prefix", "car__plate_number", "start_date", "created_at")
    )

    return render(
        request,
        "booking/admin_dashboard.html",
        {
            "total_cars": total_cars,
            "total_users": total_users,
            "active_bookings": active_bookings,
            "today": today,  # ✅ ส่งให้ template ใช้โชว์หัวข้อ “วันนี้”
        },
    )


@admin_required
def admin_employee_list(request):
    profiles = (
        Profile.objects.select_related("user")
        .annotate(
            status_rank=Case(
                When(work_status="ACTIVE", then=Value(0)),
                default=Value(1),
                output_field=IntegerField(),
            )
        )
        .order_by("status_rank", "user__username")
    )
    return render(request, "booking/admin_employees.html", {"profiles": profiles})

@admin_required
def admin_employee_create(request):
    form = EmployeeCreateForm(request.POST or None)

    if request.method == "POST" and form.is_valid():
        employee_code = form.cleaned_data["employee_code"].strip()
        first_name = form.cleaned_data["first_name"].strip()
        last_name = form.cleaned_data["last_name"].strip()
        division = form.cleaned_data["division"]
        department = form.cleaned_data["department"]
        position = form.cleaned_data["position"]
        role = form.cleaned_data["role"]

        if User.objects.filter(username=employee_code).exists():
            messages.error(request, "รหัสพนักงานนี้มีอยู่แล้ว")
        else:
            # policy: username=รหัสพนักงาน, password=รหัสพนักงาน
            user = User.objects.create_user(
                username=employee_code,
                password=employee_code,
                first_name=first_name,
                last_name=last_name,
            )
            Profile.objects.create(
                user=user,
                division=division,
                department=department,
                position=position,
                role=role,
            )
            messages.success(request, "เพิ่มพนักงานเรียบร้อย (รหัสผ่าน = รหัสพนักงาน)")
            return redirect("booking:admin_employee_list")

    return render(
        request, "booking/admin_employee_form.html", {"form": form, "mode": "create"}
    )

@admin_required
@require_POST
def admin_employee_toggle_status(request, profile_id):
    profile = get_object_or_404(Profile, id=profile_id)

    if profile.work_status == "ACTIVE":
        profile.work_status = "INACTIVE"
        msg = f"{profile.user.get_full_name()} ถูกปรับเป็นพ้นสภาพแล้ว"
    else:
        profile.work_status = "ACTIVE"
        msg = f"{profile.user.get_full_name()} กลับมาปฏิบัติงานแล้ว"

    profile.save(update_fields=["work_status"])
    messages.success(request, msg)
    return redirect("booking:admin_employee_list")


@admin_required
def admin_employee_edit(request, user_id: int):
    user = get_object_or_404(User, id=user_id)
    profile = get_profile(user)

    initial = {
        "employee_code": user.username,
        "first_name": user.first_name,
        "last_name": user.last_name,
        "division": profile.division if profile else "",
        "department": profile.department if profile else "",
        "position": profile.position if profile else "",
        "role": profile.role if profile else "EMP",
    }

    form = EmployeeUpdateForm(request.POST or None, initial=initial)

    if request.method == "POST" and form.is_valid():
        old_username = user.username
        new_username = form.cleaned_data["employee_code"].strip()

        user.username = new_username
        user.first_name = form.cleaned_data["first_name"].strip()
        user.last_name = form.cleaned_data["last_name"].strip()

        # ✅ สำคัญ: เปลี่ยนรหัสพนักงานเมื่อไร รหัสผ่านต้องเปลี่ยนตาม (policy)
        if new_username and new_username != old_username:
            user.set_password(new_username)

        user.save()

        if profile is None:
            profile = Profile.objects.create(user=user)

        profile.division = form.cleaned_data["division"]
        profile.department = form.cleaned_data["department"]
        profile.position = form.cleaned_data["position"]
        profile.role = form.cleaned_data["role"]
        profile.save()

        messages.success(request, "บันทึกข้อมูลพนักงานเรียบร้อย")
        return redirect("booking:admin_employee_list")

    return render(
        request,
        "booking/admin_employee_form.html",
        {"form": form, "mode": "edit", "user_obj": user},
    )


@admin_required
def admin_employee_delete(request, user_id):
    user = get_object_or_404(User, id=user_id)
    profile = getattr(user, "profile", None)

    if request.method == "POST":
        # ปิดการใช้งาน user
        user.is_active = False
        user.save(update_fields=["is_active"])

        # (ถ้ามี role / status)
        if profile:
            profile.role = "INACTIVE"   # หรือเก็บ field status เพิ่ม
            profile.save(update_fields=["role"])

        messages.success(request, "ปิดการใช้งานพนักงานเรียบร้อย")
    return redirect("booking:admin_employee_list")



@admin_required
def admin_car_list(request):
    cars = (
        Car.objects.all()
        .annotate(
            status_rank=Case(
                When(status="RETIRED", then=Value(1)),  # ยกเลิกใช้งานลงล่าง
                default=Value(0),
                output_field=IntegerField(),
            )
        )
        .order_by("status_rank", "plate_prefix", "plate_number")
    )
    return render(request, "booking/admin_cars.html", {"cars": cars})


@admin_required
def admin_car_create(request):
    form = CarForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "เพิ่มรถเรียบร้อย")
        return redirect("booking:admin_car_list")
    return render(
        request, "booking/admin_car_form.html", {"form": form, "mode": "create"}
    )


@admin_required
def admin_car_edit(request, car_id: int):
    car = get_object_or_404(Car, id=car_id)
    form = CarForm(request.POST or None, instance=car)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "บันทึกข้อมูลรถเรียบร้อย")
        return redirect("booking:admin_car_list")
    return render(
        request,
        "booking/admin_car_form.html",
        {"form": form, "mode": "edit", "car": car},
    )


@login_required
@admin_required
def admin_car_delete(request, car_id):
    car = get_object_or_404(Car, id=car_id)

    # ✅ ไม่ลบแล้ว เปลี่ยนเป็นยกเลิกใช้งาน
    car.status = "RETIRED"   # หรือ "OUT_OF_SERVICE"
    car.save(update_fields=["status"])

    messages.success(request, "ยกเลิกใช้งานรถคันนี้เรียบร้อยแล้ว")
    return redirect("booking:admin_car_list")  # ใช้ชื่อ url list 


# =============================================================================
# ADMIN: Booking list/detail + Monthly audit
# =============================================================================


@login_required
def admin_booking_list(request):

    # -----------------------------
    # 1) อ่านค่ากรองจาก querystring
    # -----------------------------
    today = timezone.localdate()

    year_raw = (request.GET.get("year") or str(today.year)).strip()
    month_raw = (request.GET.get("month") or str(today.month)).strip()

    try:
        year = int(year_raw)
    except ValueError:
        year = today.year

    try:
        month = int(month_raw)
    except ValueError:
        month = today.month

    # กันหลุดช่วง
    if month < 1 or month > 12:
        month = today.month

    car_id_raw = (request.GET.get("car_id") or "").strip()
    selected_car_id = int(car_id_raw) if car_id_raw.isdigit() else None

    status_filter = (request.GET.get("status") or "all").strip().lower()
    if status_filter not in ("all", "book", "return"):
        status_filter = "all"

    # -----------------------------
    # 2) ช่วงวันของเดือนที่เลือก
    # -----------------------------
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    # -----------------------------
    # 3) dropdown รถทั้งหมด
    # -----------------------------
    cars = Car.objects.all().order_by("plate_prefix", "plate_number")

    # -----------------------------
    # 4) Query หลัก: bookings ตามเดือน/ปี
    # -----------------------------
    # เลือกกรองเดือนจาก end_date (เข้าใจง่ายสุด: ทริปที่ "จบ" ในเดือนนั้น)
    # ถ้าเตงอยากใช้ updated_at (คืนจริง) ให้เปลี่ยน filter เป็น updated_at__date
    qs = (
        Booking.objects.select_related("car", "requester")
        .filter(end_date__gte=first_day, end_date__lte=last_day)
        .order_by("-end_date", "-id")
    )

    if selected_car_id:
        qs = qs.filter(car_id=selected_car_id)

    # -----------------------------
    # 5) logic ปุ่ม 3 อัน (ทั้งหมด/เฉพาะการจอง/เฉพาะคืนรถ)
    # -----------------------------
    # NOTE: สถานะในระบบเตงใช้: BOOKED, IN_USE, RETURNED, CANCELLED
    if status_filter == "book":
        qs = qs.filter(status__in=["BOOKED", "IN_USE"])
    elif status_filter == "return":
        qs = qs.filter(status="RETURNED")
    # all = ไม่กรองเพิ่ม (รวมทุกสถานะในเดือน)

    bookings = list(qs)

    # -----------------------------
    # 6) Summary (นับตามเดือน/ปี/รถที่เลือก)
    # -----------------------------
    base = Booking.objects.filter(end_date__gte=first_day, end_date__lte=last_day)
    if selected_car_id:
        base = base.filter(car_id=selected_car_id)

    count_returned = base.filter(status="RETURNED").count()
    count_in_use = base.filter(status="IN_USE").count()
    count_all = base.count()

    # -----------------------------
    # 7) months list สำหรับ dropdown
    # -----------------------------
    months = [
        {"value": 1, "label": "มกราคม"},
        {"value": 2, "label": "กุมภาพันธ์"},
        {"value": 3, "label": "มีนาคม"},
        {"value": 4, "label": "เมษายน"},
        {"value": 5, "label": "พฤษภาคม"},
        {"value": 6, "label": "มิถุนายน"},
        {"value": 7, "label": "กรกฎาคม"},
        {"value": 8, "label": "สิงหาคม"},
        {"value": 9, "label": "กันยายน"},
        {"value": 10, "label": "ตุลาคม"},
        {"value": 11, "label": "พฤศจิกายน"},
        {"value": 12, "label": "ธันวาคม"},
    ]

    context = {
        # ตาราง
        "bookings": bookings,

        # tabs เดิม
        "status_filter": status_filter,

        # filter bar ใหม่
        "months": months,
        "month": month,
        "year": year,
        "cars": cars,
        "selected_car_id": selected_car_id,

        # summary ใหม่
        "count_returned": count_returned,
        "count_in_use": count_in_use,
        "count_all": count_all,
    }

    return render(request, "booking/admin_booking_list.html", context)
@admin_required
def admin_booking_detail(request, booking_id: int):
    booking = get_object_or_404(
        Booking.objects.select_related(
            "car", "requester", "returned_by"
        ).prefetch_related("co_travelers", "co_travelers__user"),
        id=booking_id,
    )
    fuels = (
        FuelRefill.objects.filter(booking=booking)
        .select_related("car")
        .order_by("-created_at")
    )

    return render(
        request,
        "booking/admin_booking_detail.html",
        {"booking": booking, "fuels": fuels},
    )


@admin_required
def admin_booking_edit(request, booking_id: int):
    booking = get_object_or_404(
        Booking.objects.select_related(
            "car", "requester", "returned_by"
        ).prefetch_related("co_travelers", "co_travelers__user"),
        id=booking_id,
    )

    cars = Car.objects.all().order_by("plate_prefix", "plate_number")
    users = User.objects.all().order_by("first_name", "last_name", "username")

    # ✅ เพิ่ม: ดึงรายการเติมน้ำมันของ booking นี้
    fuels = (
        FuelRefill.objects.filter(booking=booking)
        .select_related("car")
        .order_by("-refill_date", "-created_at")
    )

    if request.method == "POST":
        booking.car_id = request.POST.get("car") or booking.car_id
        booking.start_date = parse_date(request.POST.get("start_date"))
        booking.end_date = parse_date(request.POST.get("end_date"))
        booking.destination = (request.POST.get("destination") or "").strip()

        # requester
        requester_id = request.POST.get("requester")
        if requester_id and str(requester_id).isdigit():
            booking.requester_id = int(requester_id)

        # returned_by (ผู้ส่งคืนรถ)
        returned_by_id = request.POST.get("returned_by") or ""
        if returned_by_id and returned_by_id.isdigit():
            booking.returned_by_id = int(returned_by_id)
        else:
            booking.returned_by = None

        # co_travelers (ผู้เดินทาง)
        co_ids = request.POST.getlist("co_travelers") or []
        co_ids = [int(x) for x in co_ids if str(x).isdigit()]
        booking.co_travelers.set(co_ids)

        # odometer
        try:
            booking.odometer_before = (
                int(request.POST.get("odometer_before"))
                if request.POST.get("odometer_before")
                else None
            )
        except (TypeError, ValueError):
            booking.odometer_before = None

        try:
            booking.odometer_after = (
                int(request.POST.get("odometer_after"))
                if request.POST.get("odometer_after")
                else None
            )
        except (TypeError, ValueError):
            booking.odometer_after = None

        booking.save()
        messages.success(request, "แก้ไขข้อมูลการยืมรถเรียบร้อยแล้ว")

        next_url = (request.POST.get("next") or request.GET.get("next") or "").strip()
        if next_url:
            return redirect(next_url)

        return redirect("booking:admin_booking_detail", booking_id=booking.id)

    return render(
        request,
        "booking/admin_booking_edit.html",
        {
            "booking": booking,
            "cars": cars,
            "users": users,
            "profiles": Profile.objects.select_related("user"),
            "fuels": fuels,  # ✅ เพิ่ม
        },
    )
@login_required
@admin_required
def admin_booking_detail_api(request, booking_id: int):
    b = get_object_or_404(
    Booking.objects.select_related(
        "car", "requester", "returned_by"
    ).prefetch_related(
        "fuel_refills",
        "co_travelers__user",
    ),
    id=booking_id,
)


    # -------------------------
    # เติมน้ำมัน
    # -------------------------
    fuel_cnt = b.fuel_refills.count()
    fuel_text = f"มี {fuel_cnt} รายการ" if fuel_cnt > 0 else "ไม่มีการเติมน้ำมัน"

    # -------------------------
    # ผู้เดินทาง
    # -------------------------
    co_names = []
    for p in b.co_travelers.all():
        u = getattr(p, "user", None)
        if u:
            co_names.append(u.get_full_name() or u.username)

    co_travelers_text = ", ".join(co_names) if co_names else "-"

    return JsonResponse(
        {
            "car": f"{b.car.plate_prefix} {b.car.plate_number} {b.car.province_full}",
            "date": f"{b.start_date:%d/%m/%Y} – {b.end_date:%d/%m/%Y}",
            "user": b.requester.get_full_name() or b.requester.username,
            "returned_by": (b.returned_by.get_full_name() if b.returned_by else None),
            "odometer_before": b.odometer_before,
            "odometer_after": b.odometer_after,
            "fuel": fuel_text,

            # ✅ เพิ่มใหม่
            "co_travelers": co_names,
            "co_travelers_text": co_travelers_text,
        }
    )


@admin_required
def admin_monthly_audit(request):
    """หน้า 'ตรวจข้อมูลรายเดือน'
    - การ์ดสรุป: คืนแล้ว / กำลังใช้งาน / ทั้งหมด (ตามเดือนที่เลือก)
    - ตาราง 2 ส่วน: (1) จอง/กำลังใช้งาน (2) คืนรถแล้ว + ตรวจเลขไมล์ + เติมน้ำมัน
    """
    today = timezone.localdate()
    year = int(request.GET.get("year") or today.year)
    month = int(request.GET.get("month") or today.month)

    tab = (request.GET.get("tab") or "audit").strip()   # audit | list (เผื่อใช้)
    kind = (request.GET.get("kind") or "all").strip()   # all | booked | returned

    car_id_raw = (request.GET.get("car_id") or "").strip()
    selected_car_id = int(car_id_raw) if car_id_raw.isdigit() else None

    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    # dropdown รถทั้งหมด
    cars = Car.objects.all().order_by("plate_prefix", "plate_number")

    # =========================================================
    # (A) คืนรถแล้ว 'ในเดือนที่เลือก' (อิงวันที่คืนจริง: updated_at)
    # =========================================================
    qs_returned_month = (
        Booking.objects.filter(status="RETURNED")
        .filter(updated_at__date__gte=first_day, updated_at__date__lte=last_day)
        .select_related("car", "requester", "returned_by")
        .prefetch_related("fuel_refills")
        .order_by("car__plate_prefix", "car__plate_number", "start_date", "end_date")
    )
    if selected_car_id:
        qs_returned_month = qs_returned_month.filter(car_id=selected_car_id)

    rows = []
    for b in qs_returned_month:
        dist = None
        issue = None

        # ✅ เช็กเลขไมล์
        if getattr(b, "odometer_before", None) is None or getattr(b, "odometer_after", None) is None:
            issue = "MISSING"
        else:
            try:
                dist = int(b.odometer_after) - int(b.odometer_before)
                if dist < 0:
                    issue = "NEGATIVE"
            except Exception:
                issue = "MISSING"

        fuel_count = 0
        if hasattr(b, "fuel_refills"):
            fuel_count = b.fuel_refills.all().count()

        rows.append({"b": b, "dist": dist, "issue": issue, "fuel_count": fuel_count})

    # =========================================================
    # (B) จอง/กำลังใช้งาน "ภายในเดือนที่เลือก" (ช่วงทับซ้อนเดือน)
    # =========================================================
    qs_booked_in_use_month = (
        Booking.objects.filter(status__in=["BOOKED", "IN_USE"])
        .filter(start_date__lte=last_day, end_date__gte=first_day)  # ทับซ้อนเดือน
        .select_related("car", "requester", "returned_by")
        .order_by("-start_date", "-id")
    )
    if selected_car_id:
        qs_booked_in_use_month = qs_booked_in_use_month.filter(car_id=selected_car_id)

    # =========================================================
    # ✅ สรุปการ์ดด้านบน
    # =========================================================
    count_returned = qs_returned_month.count()
    count_in_use = qs_booked_in_use_month.filter(status="IN_USE").count()
    count_all = count_returned + qs_booked_in_use_month.count()

    # =========================================================
    # เลือกข้อมูลตามปุ่ม (all / booked / returned)
    # =========================================================
    bookings_list = qs_booked_in_use_month
    if kind == "booked":
        rows_to_show = []
    elif kind == "returned":
        bookings_list = Booking.objects.none()
        rows_to_show = rows
    else:
        rows_to_show = rows

    return render(
        request,
        "booking/admin_monthly_audit.html",
        {
            "tab": tab,
            "kind": kind,
            "year": year,
            "month": month,
            "cars": cars,
            "selected_car_id": selected_car_id,
            "rows": rows_to_show,
            "bookings_list": bookings_list[:300],
            "count_returned": count_returned,
            "count_in_use": count_in_use,
            "count_all": count_all,
        },
    )

@login_required
@user_passes_test(is_admin)
def admin_export_fuel_excel(request):
    month = int(request.GET.get("month", timezone.localdate().month))
    year = int(request.GET.get("year", timezone.localdate().year))
    car_id = request.GET.get("car_id")

    template_excel = os.path.join(
        settings.BASE_DIR, "booking", "report_templates", "fuel_report_template.xlsx"
    )

    # ===== เลือกคันเดียว =====
    if car_id and str(car_id).isdigit():
        car = Car.objects.get(id=int(car_id))
        refills = FuelRefill.objects.filter(
            car=car, refill_date__year=year, refill_date__month=month
        ).order_by("refill_date")

        car_display = (
            f"{car.plate_prefix} {car.plate_number} {car.province_full}".strip()
        )

        # --- หัวรายงาน (ตามแบบฟอร์มราชการ) ---
        title_text = f"รายงานการใช้น้ำมันเชื้อเพลิงและน้ำมันหล่อลื่น ประจำเดือน {month:02d}/{year}"

        # รหัสพาหนะ/เลขครุภัณฑ์: ถ้าไม่มี field ให้แสดงเป็น '-'
        vehicle_code = (
            getattr(car, "vehicle_code", None)
            or getattr(car, "asset_code", None)
            or getattr(car, "asset_no", None)
            or "-"
        )

        # ประเภทรถ/ยี่ห้อรุ่น (รองรับหลายชื่อ field)
        car_type = (
            " ".join(
                [
                    str(
                        getattr(car, "brand_name", "")
                        or getattr(car, "brand", "")
                        or ""
                    ).strip(),
                    str(
                        getattr(car, "model_name", "")
                        or getattr(car, "model", "")
                        or ""
                    ).strip(),
                ]
            ).strip()
            or "-"
        )

        carline_text = f"เลขทะเบียนรถ {car_display}    รหัสพาหนะ {vehicle_code}    ประเภทรถ {car_type}"

        # --- เลขไมล์ต้นเดือน: เอาจาก booking ที่ RETURNED รายการแรกของเดือนนั้น ---
        first_booking = (
            Booking.objects.filter(
                car=car,
                status="RETURNED",
                updated_at__year=year,
                updated_at__month=month,
                odometer_before__isnull=False,
            )
            .order_by("updated_at", "id")
            .first()
        )
        odo_start = int(first_booking.odometer_before) if first_booking else None

        bio = build_fuel_excel(
            template_excel, title_text, carline_text, odo_start, refills
        )

        filename = (
            f"รายงานน้ำมัน_{car.plate_prefix}{car.plate_number}_{year}_{month:02d}.xlsx"
        )
        resp = HttpResponse(
            bio.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        resp["Content-Disposition"] = f'attachment; filename="{filename}"'
        return resp

    # ===== รวมทั้งหมด: ZIP คันละไฟล์ =====
    return admin_export_fuel_excel_all_zip(month, year, template_excel)


def admin_export_fuel_excel_all_zip(month: int, year: int, template_excel: str):
    cars = Car.objects.all().order_by("plate_prefix", "plate_number")

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for car in cars:
            refills = FuelRefill.objects.filter(
                car=car, refill_date__year=year, refill_date__month=month
            ).order_by("refill_date")

            car_display = (
                f"{car.plate_prefix} {car.plate_number} {car.province_full}".strip()
            )
            title_text = (
                f"รายงานการใช้น้ำมันเชื้อเพลิงและน้ำมันหล่อลื่น ประจำเดือน {month:02d}/{year}"
            )
            carline_text = f"เลขทะเบียนรถ {car_display}"
            odo_start = None

            bio = build_fuel_excel(
                template_excel, title_text, carline_text, odo_start, refills
            )

            inner_name = f"รายงานน้ำมัน_{car.plate_prefix}{car.plate_number}_{year}_{month:02d}.xlsx"
            zf.writestr(inner_name, bio.getvalue())

    zip_buffer.seek(0)
    filename = f"รายงานน้ำมัน_รวมทั้งหมด_{year}_{month:02d}.zip"
    resp = HttpResponse(zip_buffer.getvalue(), content_type="application/zip")
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


@admin_required
def admin_fuel_refill_edit(request, refill_id: int):
    """
    แอดมินแก้ไขรายการเติมน้ำมัน
    """
    refill = get_object_or_404(
        FuelRefill.objects.select_related("car", "booking"),
        id=refill_id,
    )

    if request.method == "POST":
        refill.refill_date = (
            parse_date(request.POST.get("refill_date")) or refill.refill_date
        )
        refill.fuel_place = (request.POST.get("fuel_place") or "").strip()
        refill.yp_number = (request.POST.get("yp_number") or "").strip()

        try:
            refill.odometer = int(request.POST.get("odometer"))
        except (TypeError, ValueError):
            messages.error(request, "เลขไมล์ต้องเป็นตัวเลข 1516")
            return redirect("booking:admin_fuel_refill_edit", refill_id=refill.id)

        refill.price_per_liter = _to_decimal(request.POST.get("price_per_liter"))
        refill.liters = _to_decimal(request.POST.get("liters"))
        refill.total_price = _to_decimal(request.POST.get("total_price"))

        if not refill.fuel_place or not refill.yp_number:
            messages.error(request, "กรุณากรอกข้อมูลให้ครบ")
            return redirect("booking:admin_fuel_refill_edit", refill_id=refill.id)

        refill.save()
        messages.success(request, "บันทึกการแก้ไขรายการเติมน้ำมันเรียบร้อยแล้ว")
        return redirect("booking:admin_booking_detail", booking_id=refill.booking_id)

    return render(
        request,
        "booking/admin_fuel_refill_edit.html",
        {
            "refill": refill,
            "booking": refill.booking,
            "car": refill.car,
        },
    )


@login_required
@user_passes_test(is_admin)
def admin_export_car_docx(request):
    month = int(request.GET.get("month", timezone.localdate().month))
    year = int(request.GET.get("year", timezone.localdate().year))
    car_id = request.GET.get("car_id")

    if not (car_id and str(car_id).isdigit()):
        return HttpResponse("กรุณาเลือกทะเบียนรถก่อนดาวน์โหลด Word", status=400)

    car = get_object_or_404(Car, id=int(car_id))

    # ✅ โลโก้ PEA ตาม path ที่เตงมีจริง
    logo_path = os.path.join(
        settings.BASE_DIR,
        "booking",
        "static",
        "booking",
        "img",
        "pea_logo.png",
    )

    # =========================
    # 1) เตรียมข้อมูลพื้นฐาน
    # =========================
    profile = get_profile(request.user)
    dept_name = profile.department if profile and profile.department else "-"

    car_plate_text = f"{car.plate_prefix} {car.plate_number} {car.province_full}".strip()
    car_brand_model = (
        f"{(car.brand_name or '').strip()} {(car.model_name or '').strip()}".strip()
        or "-"
    )

    # =========================
    # 2) หาเลขไมล์เริ่มต้น/สิ้นสุด (จาก booking RETURNED ของเดือนนั้น)
    # =========================
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    month_bookings = (
        Booking.objects.filter(
            car=car,
            status="RETURNED",
            updated_at__date__gte=first_day,
            updated_at__date__lte=last_day,
        )
        .exclude(odometer_before__isnull=True, odometer_after__isnull=True)
        .order_by("updated_at", "id")
    )

    mileage_start = None
    for b in month_bookings:
        if b.odometer_before is not None:
            mileage_start = int(b.odometer_before)
            break

    mileage_end = None
    for b in month_bookings.order_by("-updated_at", "-id"):
        if b.odometer_after is not None:
            mileage_end = int(b.odometer_after)
            break

    # =========================
    # 3) เตรียม path template Word
    # =========================
    template_path = os.path.join(
        settings.BASE_DIR,
        "booking",
        "report_templates",
        "car_report_template.docx",
    )

    # -------------------------
    # เตรียมค่าเดือน/ปีไทย + สัญญา (กัน NameError)
    # -------------------------
    TH_MONTHS = [
        "",
        "มกราคม",
        "กุมภาพันธ์",
        "มีนาคม",
        "เมษายน",
        "พฤษภาคม",
        "มิถุนายน",
        "กรกฎาคม",
        "สิงหาคม",
        "กันยายน",
        "ตุลาคม",
        "พฤศจิกายน",
        "ธันวาคม",
    ]
    month_th = TH_MONTHS[month] if 1 <= month <= 12 else str(month)
    year_th = str(year + 543)

    # ถ้าอนาคตมี field จริงค่อยผูก แต่ตอนนี้ fix ให้ไม่พัง
    agreement_no = "ฉ.1กกค.(ช)03/2566"
    agreement_date = "30 พ.ค. 2566"

    # -------------------------
    # ค่าอื่น ๆ ที่ template ต้องใช้ (กัน NameError)
    # -------------------------
    station = ""  # ถ้า profile มีฟิลด์สถานีค่อยผูก

    brand = (getattr(car, "brand_name", "") or getattr(car, "brand", "") or "").strip()
    model = (getattr(car, "model_name", "") or getattr(car, "model", "") or "").strip()

    # คนลงชื่อ/ตำแหน่ง (กัน NameError: sign_name)
    sign_name = request.user.get_full_name() or request.user.username
    sign_position = ""
    if profile and getattr(profile, "position", None):
        sign_position = profile.position

    # -------------------------
    # ✅ mapping ต้องใช้ token แบบ {{...}} ให้ตรง template จริง
    # -------------------------
    mapping = {
        "{{AGREEMENT_NO}}": agreement_no,
        "{{AGREEMENT_DATE}}": agreement_date,
        "{{MONTH_TH}}": month_th,
        "{{YEAR_TH}}": year_th,
        "{{DEPT_NAME}}": dept_name,
        "{{STATION}}": station,
        "{{BRAND}}": brand,
        "{{MODEL}}": model,
        "{{PLATE}}": car_plate_text,
        "{{MILEAGE_START}}": f"{mileage_start:,}" if mileage_start is not None else "",
        "{{MILEAGE_END}}": f"{mileage_end:,}" if mileage_end is not None else "",
        "{{SIGN_NAME}}": sign_name,
        "{{SIGN_POSITION}}": sign_position,
        "{{SIGN_DATE}}": "",  # เว้นให้เขียนมือ
    }

    # -------------------------
    # สร้าง docx จาก template
    # -------------------------
    bio = build_car_docx(template_path, mapping)

    # ✅ ตั้งชื่อไฟล์ (มีทะเบียน + จังหวัด)
    plate_compact = f"{car.plate_prefix}{car.plate_number}".strip()
    province = (car.province_full or "").strip()

    filename = (
        f"กดส.ฉ.1_รายงานการใช้น้ำมันเชื้อเพลิง_ทะเบียน_{plate_compact}_{province}"
        f"_ประจำเดือน_{month:02d}_{year}.docx"
    )

    resp = HttpResponse(
        bio.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    # ✅ บรรทัดสำคัญ อยู่ตรงนี้เท่านั้น
    resp["Content-Disposition"] = (
        f'attachment; filename="{filename}"; filename*=UTF-8\'\'{quote(filename)}'
    )

    return resp


@admin_required
@require_POST
def admin_fuel_refill_update_ajax(request, refill_id: int):
    """อัปเดต FuelRefill แบบ inline (AJAX)

    รองรับทั้ง 2 รูปแบบ:
    - fetch ส่ง JSON (Content-Type: application/json)  ✅ ใช้ใน admin_monthly_audit.html
    - form POST ปกติ (application/x-www-form-urlencoded / multipart)  ✅ เผื่อใช้ที่อื่น
    """
    refill = get_object_or_404(FuelRefill, id=refill_id)

    # -----------------------------
    # 1) รับ payload ให้ได้ทั้ง JSON และ form
    # -----------------------------
    payload = None
    ct = (request.content_type or "").lower()

    if "application/json" in ct:
        try:
            payload = json.loads(request.body.decode("utf-8") or "{}")
        except Exception:
            return JsonResponse(
                {"ok": False, "error": "รูปแบบข้อมูลไม่ถูกต้อง (JSON)"},
                status=400,
            )
    else:
        payload = request.POST

    def _get(key: str) -> str:
        val = payload.get(key) if hasattr(payload, "get") else None
        if val is None:
            return ""
        return str(val).strip()

    # -----------------------------
    # 2) mapping + validate
    # -----------------------------
    refill.fuel_place = _get("fuel_place")
    refill.yp_number = _get("yp_number")

    odo_raw = _get("odometer")
    if not odo_raw:
        return JsonResponse({"ok": False, "error": "กรุณากรอกเลขไมล์"}, status=400)
    try:
        refill.odometer = int(odo_raw)
    except (TypeError, ValueError):
        return JsonResponse({"ok": False, "error": "เลขไมล์ต้องเป็นตัวเลข"}, status=400)

    # decimal (ยอมให้ว่างได้ -> None)
    refill.price_per_liter = _to_decimal(_get("price_per_liter"))
    refill.liters = _to_decimal(_get("liters"))
    refill.total_price = _to_decimal(_get("total_price"))

    # วันที่ (ยอมให้ว่างได้)
    raw_date = _get("refill_date")
    if raw_date:
        parsed = parse_date(raw_date)
        if not parsed:
            return JsonResponse({"ok": False, "error": "รูปแบบวันที่ไม่ถูกต้อง"}, status=400)
        refill.refill_date = parsed

    if not refill.fuel_place:
        return JsonResponse({"ok": False, "error": "กรุณากรอกสถานที่เติมน้ำมัน"}, status=400)
    if not refill.yp_number:
        return JsonResponse({"ok": False, "error": "กรุณากรอกเลขยพ"}, status=400)

    # -----------------------------
    # 3) save + response
    # -----------------------------
    refill.save()

    refill_date_th = "-"
    if refill.refill_date:
        refill_date_th = refill.refill_date.strftime("%d/%m/%Y")

    return JsonResponse(
        {
            "ok": True,
            "fuel_place": refill.fuel_place or "-",
            "yp_number": refill.yp_number or "-",
            "odometer": refill.odometer if refill.odometer is not None else "-",
            "price_per_liter": str(refill.price_per_liter) if refill.price_per_liter is not None else "-",
            "liters": str(refill.liters) if refill.liters is not None else "-",
            "total_price": str(refill.total_price) if refill.total_price is not None else "-",
            "refill_date_th": refill_date_th,
        }
<<<<<<< HEAD
    )
=======
    )
>>>>>>> 1dd83bc (แกไขการแกไขเตมนำมนเพมสถานะพนกงานพนสภาพ และจดเรยงพรอมแกไขรายการรถ/พนกงาน)
