# booking/admin.py
from django.contrib import admin
from .models import Profile, Car, Booking, FuelRefill


@admin.register(Profile)
class ProfileAdmin(admin.ModelAdmin):
    list_display = ("user", "department", "position", "role")
    list_filter = ("role", "department")
    search_fields = (
        "user__username",
        "user__first_name",
        "user__last_name",
        "department",
        "position",
    )


@admin.register(Car)
class CarAdmin(admin.ModelAdmin):
    list_display = (
        "display_plate",
        "brand_name",
        "model_name",
        "usage_type",
        "status",
        "seat_count",
        "current_odometer",
    )
    list_filter = ("usage_type", "status", "gear_type")
    search_fields = (
        "plate_prefix",
        "plate_number",
        "province_full",
        "brand_name",
        "model_name",
    )


@admin.register(Booking)
class BookingAdmin(admin.ModelAdmin):
    list_display = ("car", "requester", "start_date", "end_date", "status")
    list_filter = ("status", "car", "start_date", "end_date")
    search_fields = (
        "car__plate_prefix",
        "car__plate_number",
        "requester__username",
        "requester__first_name",
        "requester__last_name",
        "destination",
    )


@admin.register(FuelRefill)
class FuelRefillAdmin(admin.ModelAdmin):
    list_display = ("car", "refill_date", "liters", "total_price", "odometer")
    list_filter = ("refill_date", "car")
    search_fields = (
        "car__plate_prefix",
        "car__plate_number",
    )
