import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('portal', '0004_marksheettemplate'),
    ]

    operations = [
        migrations.CreateModel(
            name='FacultyDutyAssignment',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('exam_type', models.CharField(choices=[('IPE', 'IPE'), ('GP', 'GP')], max_length=5)),
                ('batch', models.CharField(help_text='Batch division e.g. A1', max_length=20)),
                ('duty_date', models.DateField()),
                ('is_active', models.BooleanField(default=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('assigned_by', models.ForeignKey(null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='duties_assigned', to=settings.AUTH_USER_MODEL)),
                ('department', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='duty_assignments', to='portal.department')),
                ('faculty', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='duty_assignments', to='portal.faculty')),
                ('subject', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='duty_assignments', to='portal.subject')),
            ],
            options={
                'ordering': ['-duty_date', 'batch'],
                'unique_together': {('faculty', 'department', 'subject', 'exam_type', 'batch', 'duty_date')},
            },
        ),
        migrations.AddField(
            model_name='markentry',
            name='duty_assignment',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='mark_entries', to='portal.facultydutyassignment'),
        ),
        migrations.AddField(
            model_name='markentry',
            name='mark_data',
            field=models.JSONField(blank=True, default=dict),
        ),
        migrations.AlterField(
            model_name='markentry',
            name='exam_session',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='marks', to='portal.examsession'),
        ),
        migrations.AlterUniqueTogether(
            name='markentry',
            unique_together=set(),
        ),
        migrations.AddConstraint(
            model_name='markentry',
            constraint=models.UniqueConstraint(condition=models.Q(('exam_session__isnull', False)), fields=('student', 'exam_session'), name='unique_mark_per_exam_session'),
        ),
        migrations.AddConstraint(
            model_name='markentry',
            constraint=models.UniqueConstraint(condition=models.Q(('duty_assignment__isnull', False)), fields=('student', 'duty_assignment'), name='unique_mark_per_duty'),
        ),
    ]
