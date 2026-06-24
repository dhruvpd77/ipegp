from django.core.management.base import BaseCommand
from portal.models import User


class Command(BaseCommand):
    help = 'Create super admin user dvp/123'

    def handle(self, *args, **options):
        if User.objects.filter(username='dvp').exists():
            user = User.objects.get(username='dvp')
            user.set_password('123')
            user.role = User.Role.SUPER_ADMIN
            user.is_superuser = True
            user.is_staff = True
            user.save()
            self.stdout.write(self.style.SUCCESS('Updated super admin: dvp / 123'))
        else:
            User.objects.create_superuser(
                username='dvp',
                password='123',
                role=User.Role.SUPER_ADMIN,
            )
            self.stdout.write(self.style.SUCCESS('Created super admin: dvp / 123'))
