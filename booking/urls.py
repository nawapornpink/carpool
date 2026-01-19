from django.urls import path
from . import views

app_name = "booking"

urlpatterns = [
    # ========= Auth =========
    path("login/", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),

    # หลังล็อกอิน เช็ค role แล้วเด้งไป dashboard ที่เหมาะสม
    path("", views.dashboard_redirect, name="dashboard_redirect"),

    # ========= ผู้ใช้งานทั่วไป (พนักงาน) =========
    path("user/dashboard/", views.user_dashboard, name="user_dashboard"),
    path("user/calendar/", views.user_calendar, name="user_calendar"),
    path("api/user/bookings/", views.user_booking_events, name="user_booking_events"),

    # ฟอร์มจองรถ
    path("user/book/", views.user_create_booking, name="user_create_booking"),
    path("ajax/available-cars/", views.ajax_available_cars, name="ajax_available_cars"),

    # คืนรถ
    path("user/return-car/", views.user_return_car, name="user_return_car"),

    # เติมน้ำมัน
    path("user/fuel/", views.user_fuel_refill, name="user_fuel_refill"),

    # ✅ ยืนยันคืนรถ (ปิดเคส หลังเติมน้ำมันครบ)
    path("user/confirm-return/<int:booking_id>/", views.user_confirm_return, name="user_confirm_return"),

    # รายละเอียดเคสของฉัน (ดูรายละเอียด + เติมน้ำมันเป็นเคสต่อเคส)
    path("my-booking/<int:booking_id>/", views.user_booking_detail, name="user_booking_detail"),

    # ยกเลิกการจอง (ผู้ใช้ทำเองได้)
    path("booking/<int:booking_id>/cancel/", views.user_cancel_booking, name="cancel_booking"),

    # API รถที่พร้อมให้จอง (ใช้ตอนโหลดฟอร์มจอง)
    path("api/available-cars/", views.api_available_cars, name="api_available_cars"),

    # ========= ฝั่งธุรการ / แอดมินระบบจองรถ =========
    path("staff/dashboard/", views.admin_dashboard, name="admin_dashboard"),

    # จัดการพนักงาน
    path("staff/employees/", views.admin_employee_list, name="admin_employee_list"),
    path("staff/employees/add/", views.admin_employee_create, name="admin_employee_create"),
    path("staff/employees/<int:user_id>/edit/", views.admin_employee_edit, name="admin_employee_edit"),
    path("staff/employees/<int:user_id>/delete/", views.admin_employee_delete, name="admin_employee_delete"),

    # จัดการรถ
    path("staff/cars/", views.admin_car_list, name="admin_car_list"),
    path("staff/cars/add/", views.admin_car_create, name="admin_car_create"),
    path("staff/cars/<int:car_id>/edit/", views.admin_car_edit, name="admin_car_edit"),
    path("staff/cars/<int:car_id>/delete/", views.admin_car_delete, name="admin_car_delete"),

    # รายการการยืมทั้งหมด / รายละเอียดการยืม (ฝั่งธุรการ)
    path("staff/bookings/", views.admin_booking_list, name="admin_booking_list"),
    path("staff/bookings/<int:booking_id>/", views.admin_booking_detail, name="admin_booking_detail"),
    path("staff/bookings/<int:booking_id>/edit/", views.admin_booking_edit, name="admin_booking_edit"),

    # ========= รายงาน / ตรวจข้อมูลรายเดือน =========
    path("reports/audit/", views.admin_monthly_audit, name="admin_monthly_audit"),

    # ดาวน์โหลดเอกสาร
    path("reports/export-fuel-excel/", views.admin_export_fuel_excel, name="admin_export_fuel_excel"),
    path("reports/export-car-docx/", views.admin_export_car_docx, name="admin_export_car_docx"),

    # ✅ API สำหรับ Modal (ดูอย่างเดียว)
    path("staff/api/booking-detail/<int:booking_id>/", views.admin_booking_detail_api, name="admin_booking_detail_api"),

    # แก้ไขเติมน้ำมัน (ฝั่งแอดมิน)
    path("staff/fuel-refill/<int:refill_id>/edit/", views.admin_fuel_refill_edit, name="admin_fuel_refill_edit"),

    # ตัวจริงที่ใช้
    path("staff/fuel-refill/<int:refill_id>/update/", views.admin_fuel_refill_update_ajax, name="admin_fuel_refill_update_ajax"),

    # ✅ alias กันพัง (เผื่อไฟล์ไหนยังเรียกชื่อเก่า)
    path("staff/fuel-refill/<int:refill_id>/update/", views.admin_fuel_refill_update_ajax, name="admin_fuel_refill_update_api"),
]
