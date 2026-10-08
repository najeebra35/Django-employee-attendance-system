@echo off
echo =========================================
echo   AttendPro - Django Attendance System
echo   Windows Setup Script
echo =========================================
echo.

echo Step 1: Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: pip install failed. Make sure Python is installed.
    pause
    exit /b 1
)

echo.
echo Step 2: Running database migrations...
python manage.py makemigrations attendance_app
python manage.py migrate

echo.
echo Step 3: Creating admin superuser...
python manage.py shell -c "from django.contrib.auth.models import User; User.objects.filter(username='admin').exists() or User.objects.create_superuser('admin', 'admin@company.com', 'admin123')"

echo.
echo Step 4: Creating sample leave types and permissions...
python manage.py shell -c "from attendance_app.models import LeaveType, UserPermission; from django.contrib.auth.models import User; [LeaveType.objects.get_or_create(name=n, defaults={'days_allowed': d}) for n,d in [('Annual Leave',30),('Sick Leave',15),('Emergency Leave',5),('Unpaid Leave',0)]]; admin=User.objects.filter(is_superuser=True).first(); all_perms=[p[0] for p in UserPermission.PERMISSION_CHOICES]; UserPermission.objects.get_or_create(user=admin, defaults={'permissions': all_perms}) if admin else None; print('Done!')"

echo.
echo =========================================
echo   Setup Complete!
echo =========================================
echo.
echo   Admin Login:
echo   Username: admin
echo   Password: admin123
echo.
echo   Now run: python manage.py runserver
echo   Then open: http://127.0.0.1:8000
echo.
echo   IMPORTANT: Change admin password after first login!
echo =========================================
pause
