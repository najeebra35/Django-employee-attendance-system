#!/bin/bash
echo "========================================="
echo "  AttendPro - Django Attendance System"
echo "========================================="
echo ""
echo "Step 1: Installing dependencies..."
pip install -r requirements.txt

echo ""
echo "Step 2: Running database migrations..."
python manage.py makemigrations attendance_app
python manage.py migrate

echo ""
echo "Step 3: Creating admin superuser..."
echo "from django.contrib.auth.models import User; User.objects.filter(username='admin').exists() or User.objects.create_superuser('admin', 'admin@company.com', 'admin123')" | python manage.py shell

echo ""
echo "Step 4: Creating sample leave types..."
python manage.py shell << 'EOF'
from attendance_app.models import LeaveType, UserPermission
from django.contrib.auth.models import User

types = [
    ('Annual Leave', 30),
    ('Sick Leave', 15),
    ('Emergency Leave', 5),
    ('Unpaid Leave', 0),
]
for name, days in types:
    LeaveType.objects.get_or_create(name=name, defaults={'days_allowed': days})

# Ensure admin has UserPermission
admin = User.objects.filter(is_superuser=True).first()
if admin:
    all_perms = [p[0] for p in UserPermission.PERMISSION_CHOICES]
    UserPermission.objects.get_or_create(user=admin, defaults={'permissions': all_perms})

print("Setup complete!")
EOF

echo ""
echo "========================================="
echo "  Setup Complete!"
echo "========================================="
echo ""
echo "  Admin Login:"
echo "  Username: admin"
echo "  Password: admin123"
echo ""
echo "  Start server: python manage.py runserver"
echo "  Open: http://127.0.0.1:8000"
echo ""
echo "  IMPORTANT: Change admin password after first login!"
echo "========================================="
