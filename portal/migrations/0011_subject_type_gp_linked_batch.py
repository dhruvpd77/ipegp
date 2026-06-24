import uuid

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('portal', '0010_duty_marks_lock_timestamps'),
    ]

    operations = [
        migrations.AddField(
            model_name='subject',
            name='subject_type',
            field=models.CharField(
                choices=[('THEORY', 'Theory'), ('PRACTICAL', 'Practical')],
                default='THEORY',
                max_length=10,
            ),
        ),
        migrations.AddField(
            model_name='gpgroup',
            name='linked_batch_id',
            field=models.UUIDField(blank=True, db_index=True, null=True),
        ),
    ]
