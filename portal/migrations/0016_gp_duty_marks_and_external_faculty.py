# Generated manually for GP duty marks entry support

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('portal', '0015_gpdutyassignment'),
    ]

    operations = [
        migrations.AddField(
            model_name='gpdutyassignment',
            name='external_faculty',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='gp_external_duties',
                to='portal.faculty',
            ),
        ),
        migrations.AddField(
            model_name='gpdutyassignment',
            name='marks_locked',
            field=models.BooleanField(
                default=False,
                help_text='When set, faculty cannot edit marks for this GP split duty.',
            ),
        ),
        migrations.AddField(
            model_name='gpdutyassignment',
            name='marks_locked_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='gpdutyassignment',
            name='marks_locked_by',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='gp_duty_marks_locked',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
        migrations.AddField(
            model_name='gpdutyassignment',
            name='marks_saved_at',
            field=models.DateTimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='markentry',
            name='gp_duty_assignment',
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.CASCADE,
                related_name='mark_entries',
                to='portal.gpdutyassignment',
            ),
        ),
        migrations.AddConstraint(
            model_name='markentry',
            constraint=models.UniqueConstraint(
                condition=models.Q(('gp_duty_assignment__isnull', False)),
                fields=('student', 'gp_duty_assignment'),
                name='unique_mark_per_gp_duty',
            ),
        ),
    ]
