"""Generate fake GP submissions for FSD-2 and PYTHON-2 (practical bundle), leaving 3–4 pending per batch."""
import random
import uuid

from django.core.management.base import BaseCommand
from django.db import transaction
from django.db.models import Q

from portal.gp_utils import get_faculty_name_for_subject
from portal.models import (
    Department, GPGroup, GPGroupMemberDetail, ProjectCase, Student, Subject,
)

FSD_NAMES = ('FSD-2', 'FSD-2 ', 'FSD 2', 'FSD-II')
PYTHON_NAMES = ('PYTHON-2', 'PYTHON-2 ', 'PYTHON 2', 'PYTHON-II', 'PYTHON-II ')

FSD_TITLES = [
    'Frames Marketing e-shop', 'Nexuscart Electronics Store', 'NextHire.AI',
    'Stock Simulation Platform', 'Campus Event Hub', 'Smart Attendance Portal',
    'Hostel Booking System', 'Library Management Pro', 'Food Court Online',
    'Vehicle Parking Tracker', 'Alumni Connect Portal', 'Skill Match Platform',
]

PYTHON_TITLES = [
    'Automated Timetable Generator', 'Weather Alert Bot', 'Expense Splitter App',
    'Quiz Master Platform', 'Resume Builder AI', 'Fitness Tracker Pro',
    'Recipe Finder App', 'Music Playlist Manager', 'Task Flow Automator',
    'Code Snippet Vault', 'Language Learning Hub', 'Budget Planner Plus',
]


def find_practical_subjects(semester):
    subjects = list(Subject.objects.filter(
        Q(semester=semester, department__isnull=True) | Q(department__semester=semester),
        subject_type=Subject.SubjectType.PRACTICAL,
    ))
    fsd = next((s for s in subjects if s.name.strip().upper().replace(' ', '-') in ('FSD-2', 'FSD-II')), None)
    if not fsd:
        fsd = next((s for s in subjects if 'FSD' in s.name.upper()), None)
    py = next((s for s in subjects if 'PYTHON' in s.name.upper()), None)
    return fsd, py


def chunk_students(students):
    """Split students into groups of 2–3."""
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


def pick_title(subject, used_in_batch):
    pool = PYTHON_TITLES if 'PYTHON' in subject.name.upper() else FSD_TITLES
    choices = [t for t in pool if t not in used_in_batch]
    if not choices:
        choices = pool
    title = random.choice(choices)
    used_in_batch.add(title)
    return title


class Command(BaseCommand):
    help = 'Fake GP data for FSD-2 + PYTHON-2; 3–4 students per batch left pending'

    def add_arguments(self, parser):
        parser.add_argument('--department', type=str, help='Department code e.g. SY_1')
        parser.add_argument('--clear', action='store_true', help='Delete existing FSD-2/PYTHON-2 GP groups first')
        parser.add_argument('--pending-min', type=int, default=3, help='Min pending students per batch')
        parser.add_argument('--pending-max', type=int, default=4, help='Max pending students per batch')

    @transaction.atomic
    def handle(self, *args, **options):
        depts = Department.objects.select_related('semester').all()
        if options['department']:
            depts = depts.filter(name__iexact=options['department'])
        if not depts.exists():
            self.stderr.write('No department found.')
            return

        pending_min = max(1, options['pending_min'])
        pending_max = max(pending_min, options['pending_max'])
        total_groups = 0
        total_pending = 0

        for dept in depts:
            fsd, python = find_practical_subjects(dept.semester)
            if not fsd or not python:
                self.stdout.write(f'Skip {dept.name}: FSD-2 or PYTHON-2 subject not found')
                continue

            subjects = [fsd, python]
            if options['clear']:
                deleted, _ = GPGroup.objects.filter(
                    department=dept, subject__in=subjects,
                ).delete()
                self.stdout.write(f'Cleared {deleted} existing groups for {dept.name}')

            cases = list(ProjectCase.objects.filter(
                semester=dept.semester, is_active=True,
            ).filter(Q(department=dept) | Q(department__isnull=True)))
            if not cases:
                for name in ('CASE 1', 'CASE 2', 'CASE 3'):
                    cases.append(ProjectCase.objects.create(
                        name=name, semester=dept.semester, department=dept, is_active=True,
                    ))

            batches = (
                Student.objects.filter(department=dept)
                .values_list('batch', flat=True).distinct().order_by('batch')
            )

            for batch in batches:
                batch_students = list(
                    Student.objects.filter(department=dept, batch=batch).order_by('roll_no')
                )
                if len(batch_students) <= pending_max:
                    self.stdout.write(f'  Skip batch {batch}: too few students')
                    continue

                shuffled = batch_students[:]
                random.shuffle(shuffled)
                leave_n = random.randint(pending_min, min(pending_max, len(shuffled) - 2))
                pending_students = shuffled[:leave_n]
                participating = shuffled[leave_n:]
                total_pending += len(pending_students)

                member_groups = chunk_students(participating)
                used_fsd_titles = set()
                used_py_titles = set()
                group_num = 0

                for members in member_groups:
                    group_num += 1
                    leader = members[0]
                    genders = {m.pk: random.choice(['Male', 'Female', 'Male', 'Male']) for m in members}
                    g_div = gender_diversity(members, genders)
                    r_div = random.choice(['Yes', 'No'])
                    linked_id = uuid.uuid4()
                    group_id = f'{batch}_{group_num}'

                    for subject in subjects:
                        used = used_py_titles if 'PYTHON' in subject.name.upper() else used_fsd_titles
                        title = pick_title(subject, used)
                        case = random.choice(cases)
                        faculty_name = get_faculty_name_for_subject(dept, batch, subject) or 'Faculty'

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
                                'subject_faculty_initials': faculty_name,
                            },
                            linked_batch_id=linked_id,
                            is_submitted=True,
                        )
                        group.members.set(members)
                        GPGroupMemberDetail.objects.filter(group=group).delete()
                        for m in members:
                            GPGroupMemberDetail.objects.create(
                                group=group,
                                student=m,
                                gender=genders[m.pk],
                                member_data={},
                            )
                        total_groups += 1

                self.stdout.write(
                    f'{dept.name} batch {batch}: {len(member_groups)} teams submitted, '
                    f'{leave_n} pending ({", ".join(s.enrollment_no for s in pending_students[:3])}…)'
                )

        self.stdout.write(self.style.SUCCESS(
            f'Created {total_groups} GP groups ({total_groups // 2} practical bundles). '
            f'{total_pending} students left pending across batches.'
        ))
