"""Generate fake GP project submissions — same member groups across all subjects, per batch."""
import random

from django.core.management.base import BaseCommand
from django.db import transaction
from django.db.models import Q

from portal.models import (
    Department, GPGroup, GPGroupMemberDetail, ProjectCase, Student, Subject,
)

PROJECT_TITLES = [
    'Smart Campus Portal', 'E-Commerce Platform', 'Inventory Management System',
    'Online Quiz Application', 'Hospital Management', 'Library Automation',
    'Attendance Tracker', 'Event Management App', 'Food Delivery System',
    'Travel Booking Portal', 'Job Portal', 'Expense Manager',
    'Student Helpdesk', 'Parking Management', 'Hostel Management System',
    'Frames Marketing e-shop', 'Nexuscart Electronics Store', 'NextHire.AI',
    'Stock Simulation Platform', 'Automated Timetable Generator',
]

FACULTY_INITIALS = ['PHA', 'JSP', 'KRP', 'MND', 'SRK']


def chunk_students(students):
    """Split batch students into groups of 2–3."""
    students = list(students)
    random.shuffle(students)
    groups = []
    i = 0
    while i < len(students):
        remaining = len(students) - i
        if remaining == 1:
            if groups and len(groups[-1]) < 3:
                groups[-1].append(students[i])
            else:
                groups.append([students[i]])
            break
        if remaining == 4:
            groups.append(students[i:i + 2])
            groups.append(students[i + 2:i + 4])
            break
        size = 3 if remaining >= 3 and random.random() > 0.35 else 2
        if remaining == 2:
            size = 2
        groups.append(students[i:i + size])
        i += size
    return groups


def gender_diversity(members, genders):
    vals = [genders.get(m.pk, 'Male') for m in members]
    return 'Yes' if 'Male' in vals and 'Female' in vals else 'No'


class Command(BaseCommand):
    help = 'Create fake GP groups per batch (2–3 students), replicated for every subject'

    def add_arguments(self, parser):
        parser.add_argument('--department', type=str, help='Department code e.g. SY1')
        parser.add_argument('--clear', action='store_true', help='Delete existing GP groups first')

    @transaction.atomic
    def handle(self, *args, **options):
        depts = Department.objects.all()
        if options['department']:
            depts = depts.filter(name__iexact=options['department'])
        if not depts.exists():
            self.stderr.write('No department found.')
            return

        if options['clear']:
            deleted, _ = GPGroup.objects.all().delete()
            self.stdout.write(f'Cleared {deleted} GP group records.')

        total = 0
        for dept in depts:
            semester = dept.semester
            subjects = list(Subject.objects.filter(
                Q(semester=semester, department__isnull=True) | Q(department=dept)
            ))
            if not subjects:
                self.stdout.write(f'Skip {dept.name}: no subjects')
                continue

            cases = list(ProjectCase.objects.filter(semester=semester))
            if not cases:
                for name in ('CASE 1', 'CASE 2', 'CASE 3'):
                    cases.append(ProjectCase.objects.create(
                        name=name, semester=semester, department=dept, is_active=True,
                    ))

            batches = (
                Student.objects.filter(department=dept)
                .values_list('batch', flat=True).distinct().order_by('batch')
            )

            for batch in batches:
                batch_students = Student.objects.filter(
                    department=dept, batch=batch
                ).order_by('roll_no')
                if not batch_students.exists():
                    continue

                member_groups = chunk_students(batch_students)
                group_num = 0

                for members in member_groups:
                    group_num += 1
                    leader = members[0]
                    genders = {}
                    for m in members:
                        genders[m.pk] = random.choice(['Male', 'Female', 'Male', 'Male'])

                    g_div = gender_diversity(members, genders)
                    r_div = random.choice(['Yes', 'No'])
                    case = random.choice(cases)
                    group_id = f'{batch}_{group_num}'

                    for subject in subjects:
                        title = random.choice(PROJECT_TITLES)
                        if GPGroup.objects.filter(
                            department=dept, subject=subject, leader=leader,
                            name=title,
                        ).exists():
                            title = f'{title} ({subject.code or subject.name})'

                        group = GPGroup.objects.create(
                            name=title,
                            leader=leader,
                            department=dept,
                            subject=subject,
                            project_case=case,
                            gender_diversity=g_div,
                            religion_diversity=r_div,
                            group_data={
                                'group_id': group_id,
                                'subject_faculty_initials': random.choice(FACULTY_INITIALS),
                            },
                            is_submitted=True,
                        )
                        group.members.set(members)
                        for m in members:
                            GPGroupMemberDetail.objects.create(
                                group=group,
                                student=m,
                                gender=genders[m.pk],
                                member_data={},
                            )
                        total += 1

                self.stdout.write(
                    f'{dept.name} batch {batch}: {len(member_groups)} groups × {len(subjects)} subjects'
                )

        self.stdout.write(self.style.SUCCESS(f'Created {total} GP project submissions.'))
