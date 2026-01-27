from django.db import models
from django.contrib.auth.models import User


class Profile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)

    # ข้อมูลหน่วยงาน
    division = models.CharField(max_length=100, blank=True, null=True)  # กอง
    department = models.CharField(max_length=100, blank=True, null=True)  # แผนก
    position = models.CharField(max_length=100, blank=True, null=True)  # ตำแหน่ง

    ROLE_CHOICES = [
        ("EMP", "พนักงาน"),
        ("ADM", "ธุรการ"),
    ]
    role = models.CharField(max_length=3, choices=ROLE_CHOICES, default="EMP")

    WORK_STATUS_CHOICES = [
        ("ACTIVE", "ปฏิบัติงานอยู่"),
        ("INACTIVE", "พ้นสภาพ"),
    ]
    work_status = models.CharField(
        max_length=10,
        choices=WORK_STATUS_CHOICES,
        default="ACTIVE"
    )

    def __str__(self):
        return self.user.get_full_name() or self.user.username


class Car(models.Model):
    """ข้อมูลรถราชการ"""

    plate_prefix = models.CharField(max_length=10)  # หมวดอักษร เช่น นค
    plate_number = models.CharField(
        max_length=10, verbose_name="เลขทะเบียนรถ"
    )  # เลขทะเบียน เช่น 3814
    province_full = models.CharField(max_length=100)  # จังหวัด เช่น ขอนแก่น

    current_odometer = models.IntegerField(default=0)  # เลขไมล์ปัจจุบัน

    brand_name = models.CharField(max_length=100)  # ยี่ห้อ เช่น ISUZU
    model_name = models.CharField(max_length=100)  # รุ่น เช่น D-MAX

    color_code = models.CharField(max_length=20, default="#377dff")  # สีที่ใช้ในปฏิทิน

    STATUS_CHOICES = [
        ("READY", "พร้อมใช้งาน"),
        ("MAINTENANCE", "ส่งซ่อม"),
        ("OUT_OF_SERVICE", "งดใช้งาน"),
        ("RETIRED", "ยกเลิกใช้งาน"),
    ]
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="READY")

    seat_count = models.IntegerField(default=5)  # จำนวนที่นั่ง

    GEAR_CHOICES = [
        ("AUTO", "อัตโนมัติ"),
        ("MANUAL", "ธรรมดา"),
    ]
    gear_type = models.CharField(max_length=10, choices=GEAR_CHOICES, default="AUTO")

    USAGE_CHOICES = [
        ("POOL", "รถเช่ากฟฉ.1"),
        ("OFFICE", "รถประจำกอง/ฝ่าย"),
        ("INSPECT", "รถตรวจการ/ออกตรวจพื้นที่"),
        ("OTHER", "อื่น ๆ"),
    ]
    usage_type = models.CharField(max_length=20, choices=USAGE_CHOICES, default="POOL")

    def __str__(self):
        return f"{self.plate_prefix} {self.plate_number} ({self.province_full})"

    @property
    def display_plate(self):
        # ใช้ในหน้า admin_cars.html
        return f"{self.plate_prefix} {self.plate_number} {self.province_full}"


class Booking(models.Model):
    """การยืมรถแต่ละครั้ง"""

    STATUS_CHOICES = [
        ("BOOKED", "จองแล้ว"),
        ("IN_USE", "กำลังใช้งาน"),
        ("RETURNED", "คืนแล้ว"),
        ("PENDING_RETURN", "รอคืนรถ"),
        ("CANCELLED", "ยกเลิก"),
    ]

    car = models.ForeignKey(Car, on_delete=models.PROTECT, related_name="bookings")
    requester = models.ForeignKey(
        User, on_delete=models.PROTECT, related_name="car_bookings"
    )

    start_date = models.DateField()  # วันไป
    end_date = models.DateField()  # วันกลับ
    destination = models.CharField(max_length=255)  # สถานที่

    # เลขไมล์ก่อนและหลังของการยืมครั้งนี้
    odometer_before = models.IntegerField(
        null=True, blank=True, verbose_name="เลขไมล์ก่อนใช้งาน"
    )
    odometer_after = models.IntegerField(
        null=True, blank=True, verbose_name="เลขไมล์หลังใช้งาน"
    )

    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="BOOKED")

    # คนที่เป็นคนส่งคืนรถ
    returned_by = models.ForeignKey(
        User,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="returned_bookings",
        verbose_name="ผู้ส่งคืนรถ",
    )

    # ผู้ร่วมเดินทาง (Profile)
    co_travelers = models.ManyToManyField(
        Profile, blank=True, related_name="co_travel_bookings"
    )

    created_at = models.DateTimeField(auto_now_add=True, null=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.car} ({self.start_date} - {self.end_date})"


class FuelRefill(models.Model):
    """บันทึกการเติมน้ำมัน"""

    car = models.ForeignKey(Car, on_delete=models.CASCADE, related_name="fuel_refills")
    booking = models.ForeignKey(
        Booking,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="fuel_refills",
    )

    refill_date = models.DateField()  # วันที่เติม
    fuel_place = models.CharField(
        "สถานที่เติมน้ำมัน", max_length=255, blank=True, default=""
    )
    liters = models.DecimalField(
        max_digits=7, decimal_places=2, help_text="จำนวนลิตรที่เติม"
    )
    total_price = models.DecimalField(
        max_digits=9, decimal_places=2, help_text="ราคารวมที่จ่าย"
    )
    odometer = models.IntegerField(help_text="เลขไมล์ ณ ตอนเติม")

    # ✅ เลข ยพ. กรอก/โชว์เฉพาะงานเติมน้ำมัน
    yp_number = models.CharField(max_length=50, verbose_name="เลข ยพ.", blank=True)

    price_per_liter = models.DecimalField(
        max_digits=7,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="ราคาน้ำมันต่อลิตร",
    )

    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.car} @ {self.refill_date}"
