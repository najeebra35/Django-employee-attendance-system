from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('attendance_app', '0002_companysettings'),
    ]

    operations = [
        migrations.AddField(
            model_name='employee',
            name='portal_user',
            field=models.OneToOneField(
                blank=True, null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='employee_profile',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
    ]
