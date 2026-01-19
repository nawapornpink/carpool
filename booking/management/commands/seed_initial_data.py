from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from booking.models import Profile, Car


PASSWORD = "Carpool@123"
DEPARTMENT = "กดส. เขต 1 (ภาคตะวันออกเฉียงเหนือ) จังหวัดอุดรธานี"


ADMIN_USERS = [
    {
        "username": "adm001",
        "first_name": "วิภาดา",
        "last_name": "พิชิตการงาน",
        "position": "เจ้าพนักงานธุรการ",
    },
    {
        "username": "adm002",
        "first_name": "ชลธิชา",
        "last_name": "อินทรรักษ์",
        "position": "เจ้าพนักงานธุรการ",
    },
    {
        "username": "adm003",
        "first_name": "ภูริณัฐ",
        "last_name": "ศรีวงศ์",
        "position": "เจ้าพนักงานธุรการ",
    },
]

EMP_USERS = [
    {
        "username": "emp001",
        "first_name": "ธนดล",
        "last_name": "แก้วพงษ์",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp002",
        "first_name": "ชนากานต์",
        "last_name": "ทองสิมา",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp003",
        "first_name": "กิตติศักดิ์",
        "last_name": "ภูมิวัฒนา",
        "position": "ช่างเทคนิค (ออกตรวจพื้นที่)",
    },
    {
        "username": "emp004",
        "first_name": "สุวพิชญ์",
        "last_name": "นครินทร์",
        "position": "วิศวกรบริการผู้ใช้ไฟฟ้า",
    },
    {
        "username": "emp005",
        "first_name": "ศรายุทธ",
        "last_name": "มูลโคตร",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp006",
        "first_name": "ภัทรวดี",
        "last_name": "ปัญญากูล",
        "position": "เจ้าหน้าที่ภาคสนาม",
    },
    {
        "username": "emp007",
        "first_name": "ชยพล",
        "last_name": "สิทธิสุข",
        "position": "วิศวกรระบบเครือข่าย",
    },
    {
        "username": "emp008",
        "first_name": "วราภรณ์",
        "last_name": "บุญมาก",
        "position": "เจ้าหน้าที่ประสานงาน",
    },
    {
        "username": "emp009",
        "first_name": "พงษ์พิสิษฐ์",
        "last_name": "สุขสมบูรณ์",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp010",
        "first_name": "จิรดา",
        "last_name": "อินทะชัย",
        "position": "เจ้าหน้าที่ตรวจมิเตอร์",
    },
    {
        "username": "emp011",
        "first_name": "ณัฐพล",
        "last_name": "คำภักดี",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp012",
        "first_name": "อรุณพร",
        "last_name": "ศรีโคตร",
        "position": "เจ้าหน้าที่ธุรการภาคสนาม",
    },
    {
        "username": "emp013",
        "first_name": "ปณัฐดนย์",
        "last_name": "ใจมั่น",
        "position": "ช่างเทคนิค (บำรุงรักษา)",
    },
    {
        "username": "emp014",
        "first_name": "นรีกานต์",
        "last_name": "ภักดีรัตน์",
        "position": "วิศวกรไฟฟ้า",
    },
    {
        "username": "emp015",
        "first_name": "ศุภกร",
        "last_name": "บัวคำ",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp016",
        "first_name": "พิชญาพร",
        "last_name": "อินแสง",
        "position": "เจ้าหน้าที่บริการผู้ใช้ไฟ",
    },
    {
        "username": "emp017",
        "first_name": "เศรษฐพงษ์",
        "last_name": "สายสกุล",
        "position": "ช่างเทคนิค (สำรวจแนวสาย)",
    },
    {
        "username": "emp018",
        "first_name": "วริษา",
        "last_name": "ทองมาก",
        "position": "เจ้าหน้าที่ฝ่ายวางแผน",
    },
    {
        "username": "emp019",
        "first_name": "คณิน",
        "last_name": "โสภา",
        "position": "พนักงานขับรถยนต์",
    },
    {
        "username": "emp020",
        "first_name": "ธนัชพร",
        "last_name": "พูนทรัพย์",
        "position": "เจ้าหน้าที่ประสานงานโครงการ",
    },
]


CARS = [
    {
        "plate_number": "งค 3814 ขก",
        "province_full": "ขอนแก่น",
        "car_type": "รถปิกอัพ",
        "brand_model": "ISUZU D-MAX",
        "color_code": "#ff6b6b",
    },
    {
        "plate_number": "3ขพ 6247 กทม",
        "province_full": "กรุงเทพมหานคร",
        "car_type": "รถปิกอัพ",
        "brand_model": "Toyota Revo",
        "color_code": "#4dabf7",
    },
    {
        "plate_number": "ขว 6808 ขก",
        "province_full": "ขอนแก่น",
        "car_type": "รถปิกอัพ 4 ประตู",
        "brand_model": "ISUZU Double Cab",
        "color_code": "#51cf66",
    },
    {
        "plate_number": "3ขพ 6250 กทม",
        "province_full": "กรุงเทพมหานคร",
        "car_type": "รถปิกอัพ",
        "brand_model": "Toyota Revo",
        "color_code": "#845ef7",
    },
    {
        "plate_number": "งค 3806 ขก",
        "province_full": "ขอนแก่น",
        "car_type": "รถปิกอัพ",
        "brand_model": "ISUZU D-MAX",
        "color_code": "#ffa94d",
    },
    {
        "plate_number": "ขว 6800 ขก",
        "province_full": "ขอนแก่น",
        "car_type": "รถปิกอัพ 4 ประตู",
        "brand_model": "ISUZU Double Cab",
        "color_code": "#20c997",
    },
    {
        "plate_number": "1ขญ 206 กทม",
        "province_full": "กรุงเทพมหานคร",
        "car_type": "รถตรวจการ / PPV",
        "brand_model": "ISUZU MU-X",
        "color_code": "#e64980",
    },
    {
        "plate_number": "ขว 6809 ขก",
        "province_full": "ขอนแก่น",
        "car_type": "รถปิกอัพ 4 ประตู",
        "brand_model": "ISUZU Double Cab",
        "color_code": "#adb5bd",
    },
]


class Command(BaseCommand):
    help = (
        "Seed initial data: admin users, employee users, and cars for กดส. เขต 1 อุดรธานี"
    )

    def handle(self, *args, **options):
        self.stdout.write(self.style.MIGRATE_HEADING("Seeding initial data..."))

        # สร้างธุรการ
        for admin in ADMIN_USERS:
            user, created = User.objects.get_or_create(
                username=admin["username"],
                defaults={
                    "first_name": admin["first_name"],
                    "last_name": admin["last_name"],
                    "is_staff": True,
                },
            )
            if created:
                user.set_password(PASSWORD)
                user.save()
                self.stdout.write(
                    self.style.SUCCESS(f"Created admin user: {user.username}")
                )
            else:
                self.stdout.write(f"Admin user already exists: {user.username}")

            profile, _ = Profile.objects.get_or_create(
                user=user,
                defaults={
                    "department": DEPARTMENT,
                    "section": "กอง กดส.",
                    "position": admin["position"],
                    "role": "ADM",
                },
            )

        # สร้างพนักงานทั่วไป
        for emp in EMP_USERS:
            user, created = User.objects.get_or_create(
                username=emp["username"],
                defaults={
                    "first_name": emp["first_name"],
                    "last_name": emp["last_name"],
                    "is_staff": False,
                },
            )
            if created:
                user.set_password(PASSWORD)
                user.save()
                self.stdout.write(
                    self.style.SUCCESS(f"Created employee user: {user.username}")
                )
            else:
                self.stdout.write(f"Employee user already exists: {user.username}")

            profile, _ = Profile.objects.get_or_create(
                user=user,
                defaults={
                    "department": DEPARTMENT,
                    "section": "กอง กดส.",
                    "position": emp["position"],
                    "role": "EMP",
                },
            )

        # สร้างรถ
        for car_data in CARS:
            car, created = Car.objects.get_or_create(
                plate_number=car_data["plate_number"],
                defaults={
                    "province_full": car_data["province_full"],
                    "car_type": car_data["car_type"],
                    "brand_model": car_data["brand_model"],
                    "color_code": car_data["color_code"],
                    "is_active": True,
                },
            )
            if created:
                self.stdout.write(
                    self.style.SUCCESS(f"Created car: {car.plate_number}")
                )
            else:
                self.stdout.write(f"Car already exists: {car.plate_number}")

        self.stdout.write(self.style.SUCCESS("Seeding completed."))
