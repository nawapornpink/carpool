from django.contrib import admin
from django.urls import path, include
from django.contrib.auth import views as auth_views
from booking import views as booking_views
urlpatterns = [
    path('admin/', admin.site.urls),

    # รวม urls ของ app booking (ของเตงน่าจะมีอยู่แล้ว)
    path('', include(('booking.urls', 'booking'), namespace='booking')),

    # ✅ หน้า login
    path('accounts/login/',
         auth_views.LoginView.as_view(template_name='booking/login.html'),
         name='login'),

    # ✅ หน้า logout
    path(
        'accounts/logout/',
        booking_views.logout_view,name='logout'),
]