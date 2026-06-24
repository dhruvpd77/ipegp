from django.core.management.base import BaseCommand
from django.utils import timezone

from portal.models import User, Student, Faculty, GeneratedCredential


class Command(BaseCommand):
    help = 'Regenerate all student and faculty credentials with 4-digit numeric passwords'

    def handle(self, *args, **options):
        student_count = self._regenerate_students()
        faculty_count = self._regenerate_faculty()
        self.stdout.write(self.style.SUCCESS(
            f'Regenerated {student_count} student and {faculty_count} faculty credentials (4-digit passwords).'
        ))

    def _regenerate_students(self):
        count = 0
        creds_to_create = []
        for stu in Student.objects.select_related('user').iterator():
            password = Student.generate_password()
            if stu.user:
                stu.user.set_password(password)
                stu.user.save(update_fields=['password'])
                user = stu.user
            else:
                user = User.objects.create_user(
                    username=stu.enrollment_no,
                    password=password,
                    role=User.Role.STUDENT,
                    first_name=stu.name,
                )
                stu.user = user
                stu.credentials_generated = True
                stu.save(update_fields=['user', 'credentials_generated'])
            creds_to_create.append(GeneratedCredential(
                user=user, plain_password=password, generated_at=timezone.now(),
            ))
            count += 1
            if count % 50 == 0:
                self.stdout.write(f'  Students: {count}...')
        if creds_to_create:
            GeneratedCredential.objects.bulk_create(creds_to_create, batch_size=200)
        return count

    def _regenerate_faculty(self):
        count = 0
        creds_to_create = []
        for fac in Faculty.objects.select_related('user').iterator():
            password = Faculty.generate_password()
            username = fac.mentor_code or f'fac_{fac.pk}'
            if fac.user:
                if fac.user.username != username:
                    fac.user.username = username
                fac.user.set_password(password)
                fac.user.save()
                user = fac.user
            else:
                user = User.objects.create_user(
                    username=username,
                    password=password,
                    role=User.Role.FACULTY,
                    first_name=fac.name,
                )
                fac.user = user
                fac.credentials_generated = True
                fac.save(update_fields=['user', 'credentials_generated'])
            creds_to_create.append(GeneratedCredential(
                user=user, plain_password=password, generated_at=timezone.now(),
            ))
            count += 1
        if creds_to_create:
            GeneratedCredential.objects.bulk_create(creds_to_create, batch_size=200)
        return count
