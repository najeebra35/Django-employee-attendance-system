from django.db import migrations, models
import datetime


class Migration(migrations.Migration):

    dependencies = [
        ('attendance_app', '0001_initial'),
    ]

    operations = [
        migrations.CreateModel(
            name='CompanySettings',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('company_name', models.CharField(default='Your Company Name', max_length=200)),
                ('company_address', models.TextField(default='Your Address, Dubai, UAE')),
                ('company_phone', models.CharField(blank=True, default='', max_length=50)),
                ('company_email', models.EmailField(blank=True, default='')),
                ('company_website', models.URLField(blank=True, default='')),
                ('company_trn', models.CharField(blank=True, default='', max_length=50, verbose_name='TRN / Tax No.')),
                ('logo', models.ImageField(blank=True, null=True, upload_to='company/')),
                ('default_in_time', models.TimeField(default=datetime.time(7, 0))),
                ('default_out_time', models.TimeField(default=datetime.time(17, 0))),
                ('work_days', models.CharField(default='Mon-Sat', help_text='e.g. Mon-Sat, Mon-Fri', max_length=20)),
                ('weekend_day', models.CharField(choices=[('Sunday', 'Sunday'), ('Friday', 'Friday'), ('Saturday', 'Saturday'), ('Friday-Saturday', 'Fri & Sat')], default='Sunday', max_length=20)),
                ('timezone_name', models.CharField(default='Asia/Dubai', max_length=60)),
                ('date_format', models.CharField(choices=[('d-m-Y', 'DD-MM-YYYY'), ('m/d/Y', 'MM/DD/YYYY'), ('Y-m-d', 'YYYY-MM-DD')], default='d-m-Y', max_length=20)),
                ('updated_at', models.DateTimeField(auto_now=True)),
            ],
            options={
                'verbose_name': 'Company Settings',
            },
        ),
    ]
