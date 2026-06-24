import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('portal', '0003_gp_project_fields'),
    ]

    operations = [
        migrations.CreateModel(
            name='MarksheetTemplate',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=150)),
                ('exam_type', models.CharField(choices=[('IPE', 'IPE Marksheet'), ('GP', 'GP Marksheet')], max_length=5)),
                ('template_file', models.FileField(upload_to='marksheet_templates/')),
                ('schema', models.JSONField(blank=True, default=dict)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('created_by', models.ForeignKey(null=True, on_delete=django.db.models.deletion.SET_NULL, to=settings.AUTH_USER_MODEL)),
                ('department', models.ForeignKey(blank=True, help_text='Leave blank for semester-wide template (all departments).', null=True, on_delete=django.db.models.deletion.CASCADE, related_name='marksheet_templates', to='portal.department')),
                ('semester', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='marksheet_templates', to='portal.semester')),
                ('subject', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='marksheet_templates', to='portal.subject')),
            ],
            options={
                'ordering': ['-created_at'],
            },
        ),
    ]
