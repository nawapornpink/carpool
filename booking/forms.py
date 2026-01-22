# booking/forms.py
from django import forms
from django.contrib.auth.models import User

from .models import Profile, Car


# ---------- ตัวเลือกกอง/แผนก/ตำแหน่ง (ปรับได้ตามจริงภายหลัง) ----------

DIVISION_CHOICES = [
    ("", "--- เลือกกอง ---"),
    ("กดส.", "กองดิจิทัลและการสื่อสาร"),
]

DEPARTMENT_CHOICES = [
    ("", "--- เลือกแผนก ---"),
    ("แผนกดิจิทัล", "แผนกดิจิทัล"),
    ("แผนกสื่อสารองค์กร", "แผนกสื่อสารองค์กร"),
    ("แผนกวิศวกรรม", "แผนกวิศวกรรม"),
    ("อื่นๆ", "อื่น ๆ"),
]

POSITION_CHOICES = [
    ("", "--- เลือกตำแหน่ง ---"),
    ("พนักงานปฏิบัติการ", "พนักงานปฏิบัติการ"),
    ("วิศวกร", "วิศวกร"),
    ("หัวหน้าแผนก", "หัวหน้าแผนก"),
    ("ผู้อำนวยการกอง", "ผู้อำนวยการกอง"),
    ("อื่นๆ", "อื่น ๆ"),
]


# ---------- ฟอร์มจัดการพนักงาน ----------


class EmployeeCreateForm(forms.Form):
    employee_code = forms.CharField(label="รหัสประจำตัวพนักงาน", max_length=20)
    first_name = forms.CharField(label="ชื่อ", max_length=150)
    last_name = forms.CharField(label="สกุล", max_length=150)
    division = forms.ChoiceField(label="กอง", choices=DIVISION_CHOICES)
    department = forms.ChoiceField(label="แผนก", choices=DEPARTMENT_CHOICES)
    position = forms.ChoiceField(label="ตำแหน่ง", choices=POSITION_CHOICES)
    role = forms.ChoiceField(label="บทบาทในระบบ", choices=Profile.ROLE_CHOICES)


class EmployeeUpdateForm(forms.Form):
    employee_code = forms.CharField(label="รหัสประจำตัวพนักงาน", max_length=20)
    first_name = forms.CharField(label="ชื่อ", max_length=150)
    last_name = forms.CharField(label="สกุล", max_length=150)
    division = forms.ChoiceField(label="กอง", choices=DIVISION_CHOICES)
    department = forms.ChoiceField(label="แผนก", choices=DEPARTMENT_CHOICES)
    position = forms.ChoiceField(label="ตำแหน่ง", choices=POSITION_CHOICES)
    role = forms.ChoiceField(label="บทบาทในระบบ", choices=Profile.ROLE_CHOICES)


# ---------- ฟอร์มจัดการรถ ----------


class CarForm(forms.ModelForm):
    """ฟอร์มเพิ่ม / แก้ไขข้อมูลรถราชการ"""

    class Meta:
        model = Car
        fields = [
            "plate_prefix",
            "plate_number",
            "province_full",
            "current_odometer",
            "brand_name",
            "model_name",
            "color_code",
            "status",
            "seat_count",
            "gear_type",
            "usage_type",
        ]
        widgets = {
            "color_code": forms.TextInput(attrs={"type": "color"}),
        }


from .models import FuelRefill


class FuelRefillForm(forms.ModelForm):
    class Meta:
        model = FuelRefill
        fields = [
            "refill_date",
            "fuel_place",
            "yp_number",
            "odometer",
            "price_per_liter",
            "liters",
            "total_price",
        ]
        widgets = {
            "refill_date": forms.DateInput(attrs={"type": "date"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for f in self.fields.values():
            f.widget.attrs.update({"class": "form-control"})


class FuelRefillForm(forms.ModelForm):
    class Meta:
        model = FuelRefill
        fields = [
            "refill_date",
            "fuel_place",
            "price_per_liter",
            "liters",
            "total_price",
            "odometer",
            "yp_number",
        ]
        labels = {
            "refill_date": "วันที่เติมน้ำมัน",
            "fuel_place": "สถานที่เติมน้ำมัน",
            "price_per_liter": "ราคาน้ำมันต่อลิตร",
            "liters": "จำนวนลิตร",
            "total_price": "รวมราคา",
            "odometer": "เลขไมล์",
            "yp_number": "เลข ยพ.",
        }
