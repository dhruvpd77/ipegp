"""Create dummy syllabus (PDF) and phase-wise papers (Word) for testing."""
import io

from django.core.files.base import ContentFile
from django.core.management.base import BaseCommand
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from portal.models import Department, Subject, SubjectSyllabus, SubjectPaper


def _make_pdf(title, lines):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setFont('Helvetica-Bold', 16)
    c.drawString(50, 800, title)
    c.setFont('Helvetica', 11)
    y = 770
    for line in lines:
        c.drawString(50, y, line)
        y -= 18
        if y < 50:
            c.showPage()
            y = 800
    c.save()
    buf.seek(0)
    return buf.read()


def _make_rtf_doc(title, lines):
    body = '\\par '.join(lines)
    rtf = (
        r'{\rtf1\ansi\deff0'
        r'{\fonttbl{\f0 Helvetica;}}'
        r'\f0\fs28\b ' + title.replace('\\', '\\\\') + r'\b0\par\par\fs22 '
        + body.replace('\\', '\\\\')
        + '}'
    )
    return rtf.encode('utf-8')


class Command(BaseCommand):
    help = 'Generate dummy syllabus and phase-wise papers for subjects'

    def add_arguments(self, parser):
        parser.add_argument('--department', default='SY1', help='Department code e.g. SY1')
        parser.add_argument('--clear', action='store_true', help='Remove existing dummy records first')

    def handle(self, *args, **options):
        dept = Department.objects.filter(code=options['department']).first()
        if not dept:
            dept = Department.objects.first()
        if not dept:
            self.stderr.write('No department found.')
            return

        subjects = Subject.objects.filter(
            semester=dept.semester,
        ).filter(
            models_Q_department_null_or_dept(dept)
        )
        if not subjects.exists():
            subjects = Subject.objects.filter(semester=dept.semester)[:2]

        if options['clear']:
            SubjectSyllabus.objects.filter(semester=dept.semester, department__isnull=True).delete()
            SubjectPaper.objects.filter(semester=dept.semester, department__isnull=True).delete()
            self.stdout.write('Cleared existing semester-wide syllabus/papers.')

        created_s, created_p = 0, 0
        for subject in subjects:
            if not SubjectSyllabus.objects.filter(
                subject=subject, semester=dept.semester, department__isnull=True, is_active=True,
            ).exists():
                pdf = _make_pdf(
                    f'{subject.name} — Syllabus',
                    [
                        'L. J. Institute of Engineering & Technology',
                        f'Subject: {subject.name}',
                        '',
                        'Unit 1: Introduction and fundamentals',
                        'Unit 2: Core concepts and practical applications',
                        'Unit 3: Advanced topics and case studies',
                        'Unit 4: IPE/GP evaluation guidelines',
                        '',
                        '(Dummy syllabus for portal testing)',
                    ],
                )
                syl = SubjectSyllabus(
                    subject=subject,
                    semester=dept.semester,
                    department=None,
                    title=f'{subject.name} Syllabus 2026',
                )
                syl.file.save(f'syllabus_{subject.code or subject.pk}.pdf', ContentFile(pdf), save=True)
                created_s += 1

            for phase in (1, 2, 3):
                if SubjectPaper.objects.filter(
                    subject=subject, semester=dept.semester,
                    department__isnull=True, phase=phase, is_active=True,
                ).exists():
                    continue
                doc = _make_rtf_doc(
                    f'{subject.name} — IPE Phase {phase}',
                    [
                        'Internal Practical Examination',
                        f'Subject: {subject.name}',
                        f'Phase: {phase}',
                        '',
                        'Program 1: Write a program to demonstrate core concepts.',
                        'Program 2: Write a program for advanced application.',
                        '',
                        '(Dummy question paper for portal testing)',
                    ],
                )
                paper = SubjectPaper(
                    subject=subject,
                    semester=dept.semester,
                    department=None,
                    phase=phase,
                    title=f'Phase {phase} — {subject.name}',
                )
                safe = (subject.code or str(subject.pk)).replace(' ', '_')
                paper.file.save(f'paper_{safe}_phase{phase}.doc', ContentFile(doc), save=True)
                created_p += 1

        self.stdout.write(self.style.SUCCESS(
            f'Done — {created_s} syllabi, {created_p} papers for {dept.name} ({subjects.count()} subjects).'
        ))


def models_Q_department_null_or_dept(dept):
    from django.db.models import Q
    return Q(department__isnull=True) | Q(department=dept)
