from waitress import serve
from carpool_site.wsgi import application

if __name__ == "__main__":
    # ปรับ host/port ตามต้องการ (เช่น 0.0.0.0:8000 หรือ 0.0.0.0:8080)
    serve(application, host="0.0.0.0", port=8000)