import secrets

from django.contrib.auth.models import AbstractUser, BaseUserManager
from django.db import models
from django.utils import timezone


def generate_four_digit_password():
    """Generate a random 4-digit numeric password (1000–9999)."""
    return str(secrets.randbelow(9000) + 1000)


class UserManager(BaseUserManager):
    def create_user(self, username, password=None, **extra_fields):
        if not username:
            raise ValueError('Username is required')
        user = self.model(username=username, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, username, password=None, **extra_fields):
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)
        extra_fields.setdefault('role', 'SUPER_ADMIN')
        return self.create_user(username, password, **extra_fields)


class User(AbstractUser):
    objects = UserManager()
    class Role(models.TextChoices):
        SUPER_ADMIN = 'SUPER_ADMIN', 'Super Admin'
        SEMESTER_ADMIN = 'SEMESTER_ADMIN', 'Semester Admin'
        DEPARTMENT_ADMIN = 'DEPARTMENT_ADMIN', 'Department Admin'
        FACULTY = 'FACULTY', 'Faculty'
        STUDENT = 'STUDENT', 'Student'

    role = models.CharField(max_length=20, choices=Role.choices, default=Role.STUDENT)
    phone = models.CharField(max_length=15, blank=True)

    def __str__(self):
        return f"{self.username} ({self.get_role_display()})"


class Semester(models.Model):
    name = models.CharField(max_length=100)
    academic_year = models.CharField(max_length=20)
    is_active = models.BooleanField(default=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='semesters_created')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.name} ({self.academic_year})"


class SemesterAdminAssignment(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='semester_admin_assignment')
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='admins')
    assigned_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.user.username} - {self.semester.name}"


class Department(models.Model):
    name = models.CharField(max_length=100)
    code = models.CharField(max_length=20, blank=True)
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='departments')
    sheet_semester_label = models.CharField(
        max_length=150, blank=True,
        help_text='Displayed on attendance sheet e.g. III- ODD-2025',
    )
    sheet_department_label = models.CharField(
        max_length=150, blank=True,
        help_text='Displayed on attendance sheet e.g. SY- CE\\IT- 1',
    )
    gp_submission_deadline = models.DateTimeField(
        null=True, blank=True,
        help_text='After this date and time students cannot submit or edit GP projects.',
    )
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='departments_created')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ['name', 'semester']
        ordering = ['name']

    def __str__(self):
        return f"{self.name} - {self.semester.name}"


class DepartmentAdminAssignment(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='department_admin_assignment')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='admins')
    assigned_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.user.username} - {self.department.name}"


class Subject(models.Model):
    class SubjectType(models.TextChoices):
        THEORY = 'THEORY', 'Theory'
        PRACTICAL = 'PRACTICAL', 'Practical'

    name = models.CharField(max_length=150)
    code = models.CharField(max_length=20, blank=True)
    subject_type = models.CharField(
        max_length=10,
        choices=SubjectType.choices,
        default=SubjectType.THEORY,
    )
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='subjects')
    department = models.ForeignKey(
        Department, on_delete=models.CASCADE, null=True, blank=True,
        related_name='subjects',
        help_text='Leave blank for semester-wide subject'
    )
    max_marks_ipe = models.PositiveIntegerField(default=25)
    max_marks_gp = models.PositiveIntegerField(default=50)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        scope = self.department.name if self.department else 'All Departments'
        return f"{self.name} ({scope})"


class Faculty(models.Model):
    name = models.CharField(max_length=150)
    mentor_code = models.CharField(max_length=50, blank=True)
    email = models.EmailField(blank=True)
    phone = models.CharField(max_length=15, blank=True)
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='faculty_members')
    is_external = models.BooleanField(
        default=False,
        help_text='External examiner — may share login if matched to existing faculty.',
    )
    user = models.OneToOneField(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='faculty_profile')
    credentials_generated = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name_plural = 'Faculty'
        ordering = ['name']

    def __str__(self):
        return self.name

    @staticmethod
    def generate_password():
        return generate_four_digit_password()


class Student(models.Model):
    roll_no = models.CharField(max_length=20)
    name = models.CharField(max_length=150)
    enrollment_no = models.CharField(max_length=20, unique=True)
    branch = models.CharField(max_length=50)
    batch = models.CharField(max_length=20)
    mentor = models.CharField(max_length=100, blank=True)
    student_phone = models.CharField(max_length=15, blank=True)
    parent_contact = models.CharField(max_length=15, blank=True)
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='students')
    user = models.OneToOneField(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='student_profile')
    credentials_generated = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['batch', 'roll_no']

    def __str__(self):
        return f"{self.name} ({self.enrollment_no})"

    @staticmethod
    def generate_password():
        return generate_four_digit_password()


class MarksheetTemplate(models.Model):
    class ExamType(models.TextChoices):
        IPE = 'IPE', 'IPE Marksheet'
        GP = 'GP', 'GP Marksheet'

    name = models.CharField(max_length=150)
    exam_type = models.CharField(max_length=5, choices=ExamType.choices)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='marksheet_templates')
    semester = models.ForeignKey(
        Semester, on_delete=models.CASCADE, null=True, blank=True,
        related_name='marksheet_templates',
    )
    department = models.ForeignKey(
        Department, on_delete=models.CASCADE, null=True, blank=True,
        related_name='marksheet_templates',
        help_text='Leave blank for semester-wide template (all departments).',
    )
    template_file = models.FileField(upload_to='marksheet_templates/')
    schema = models.JSONField(default=dict, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        scope = self.department.name if self.department else 'All Departments'
        return f'{self.subject.name} - {self.exam_type} ({scope})'


class AttendanceSheetTemplate(models.Model):
    class ExamType(models.TextChoices):
        IPE = 'IPE', 'IPE Attendance Sheet'
        GP = 'GP', 'GP Attendance Sheet'

    name = models.CharField(max_length=150)
    exam_type = models.CharField(max_length=5, choices=ExamType.choices)
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='attendance_templates', null=True, blank=True)
    department = models.ForeignKey(Department, on_delete=models.CASCADE, null=True, blank=True, related_name='attendance_templates')
    template_file = models.FileField(upload_to='attendance_templates/')
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'{self.name} ({self.exam_type})'


class SubjectSyllabus(models.Model):
    """Subject-wise syllabus document (PDF / Word) for faculty."""
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='syllabi')
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='syllabi')
    department = models.ForeignKey(
        Department, on_delete=models.CASCADE, null=True, blank=True,
        related_name='syllabi',
        help_text='Leave blank for semester-wide (all departments).',
    )
    title = models.CharField(max_length=200, blank=True)
    file = models.FileField(upload_to='syllabus/')
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['subject__name', '-created_at']
        verbose_name_plural = 'Subject syllabi'

    def __str__(self):
        return f'{self.subject.name} Syllabus'

    @property
    def display_title(self):
        return self.title or f'{self.subject.name} Syllabus'

    @property
    def file_ext(self):
        name = (self.file.name or '').lower()
        if name.endswith('.pdf'):
            return 'PDF'
        if name.endswith('.docx'):
            return 'DOCX'
        if name.endswith('.doc'):
            return 'DOC'
        return 'FILE'


class SubjectPaper(models.Model):
    """Phase-wise IPE/GP question papers (PDF / Word) per subject."""
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='papers')
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='papers')
    department = models.ForeignKey(
        Department, on_delete=models.CASCADE, null=True, blank=True,
        related_name='papers',
        help_text='Leave blank for semester-wide (all departments).',
    )
    phase = models.PositiveSmallIntegerField(help_text='Paper phase e.g. 1, 2, 3')
    title = models.CharField(max_length=200, blank=True)
    file = models.FileField(upload_to='papers/')
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['subject__name', 'phase']
        unique_together = ['subject', 'semester', 'department', 'phase']

    def __str__(self):
        return f'{self.subject.name} — Phase {self.phase}'

    @property
    def display_title(self):
        return self.title or f'Phase {self.phase} — {self.subject.name}'

    @property
    def file_ext(self):
        name = (self.file.name or '').lower()
        if name.endswith('.pdf'):
            return 'PDF'
        if name.endswith('.docx'):
            return 'DOCX'
        if name.endswith('.doc'):
            return 'DOC'
        return 'FILE'


class FormTemplate(models.Model):
    name = models.CharField(max_length=150)
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='form_templates')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, null=True, blank=True, related_name='form_templates')
    exam_type = models.CharField(max_length=10, choices=[('IPE', 'IPE'), ('GP', 'GP')], default='GP')
    uploaded_file = models.FileField(upload_to='form_templates/', blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.name} ({self.exam_type})"


class FormField(models.Model):
    class FieldScope(models.TextChoices):
        GROUP = 'GROUP', 'Group Information'
        MEMBER = 'MEMBER', 'Individual / Member Information'

    class FieldType(models.TextChoices):
        TEXT = 'text', 'Text'
        TEXTAREA = 'textarea', 'Textarea'
        NUMBER = 'number', 'Number'
        EMAIL = 'email', 'Email'
        YES_NO = 'yes_no', 'Yes / No'
        SELECT = 'select', 'Dropdown Select'

    template = models.ForeignKey(FormTemplate, on_delete=models.CASCADE, related_name='fields')
    field_name = models.CharField(max_length=150)
    field_label = models.CharField(max_length=200)
    field_scope = models.CharField(max_length=10, choices=FieldScope.choices, default=FieldScope.GROUP)
    field_type = models.CharField(max_length=20, choices=FieldType.choices, default=FieldType.TEXT)
    choices = models.TextField(blank=True, help_text='Comma-separated options for dropdown fields')
    is_required = models.BooleanField(default=True)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['field_scope', 'order']

    def __str__(self):
        return f"{self.field_label} ({self.get_field_scope_display()})"

    def choice_list(self):
        return [c.strip() for c in self.choices.split(',') if c.strip()]


class ProjectCase(models.Model):
    """Selectable cases for GP project submissions."""
    name = models.CharField(max_length=150)
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='project_cases')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, null=True, blank=True, related_name='project_cases')
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class GPGroup(models.Model):
    name = models.CharField(max_length=150, verbose_name='Project Title')
    leader = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='led_groups')
    members = models.ManyToManyField(Student, related_name='gp_groups')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='gp_groups')
    subject = models.ForeignKey(Subject, on_delete=models.SET_NULL, null=True, blank=True)
    project_case = models.ForeignKey(ProjectCase, on_delete=models.SET_NULL, null=True, blank=True, related_name='groups')
    gender_diversity = models.CharField(max_length=5, blank=True, choices=[('Yes', 'Yes'), ('No', 'No')])
    religion_diversity = models.CharField(max_length=5, blank=True, choices=[('Yes', 'Yes'), ('No', 'No')])
    group_data = models.JSONField(default=dict, blank=True)
    linked_batch_id = models.UUIDField(null=True, blank=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    is_submitted = models.BooleanField(default=False)

    def __str__(self):
        return self.name


class GPGroupMemberDetail(models.Model):
    group = models.ForeignKey(GPGroup, on_delete=models.CASCADE, related_name='member_details')
    student = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='gp_member_details')
    gender = models.CharField(max_length=20, blank=True)
    region = models.CharField(max_length=100, blank=True, verbose_name='Area name')
    religion = models.CharField(max_length=50, blank=True)
    member_data = models.JSONField(default=dict, blank=True)

    class Meta:
        unique_together = [('group', 'student')]
        ordering = ['student__roll_no']

    def __str__(self):
        return f"{self.student.name} in {self.group.name}"


class FormSubmission(models.Model):
    group = models.ForeignKey(GPGroup, on_delete=models.CASCADE, null=True, blank=True, related_name='submissions')
    student = models.ForeignKey(Student, on_delete=models.CASCADE, null=True, blank=True, related_name='form_submissions')
    template = models.ForeignKey(FormTemplate, on_delete=models.CASCADE, related_name='submissions')
    field_data = models.JSONField(default=dict)
    submitted_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        owner = self.group.name if self.group else self.student.name
        return f"{owner} - {self.template.name}"


class ExternalRegistrationForm(models.Model):
    """Public external examiner registration form — scoped to a semester."""
    title = models.CharField(max_length=200, default='External Examiner Registration')
    semester = models.ForeignKey(Semester, on_delete=models.CASCADE, related_name='external_forms')
    description = models.TextField(blank=True)
    is_active = models.BooleanField(default=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='external_forms_created')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        status = 'Active' if self.is_active else 'Inactive'
        return f'{self.title} — {self.semester.name} ({status})'


class ExternalRegistrationField(models.Model):
    class FieldType(models.TextChoices):
        TEXT = 'text', 'Text'
        TEXTAREA = 'textarea', 'Textarea'
        NUMBER = 'number', 'Number'
        EMAIL = 'email', 'Email'
        DATE = 'date', 'Date'
        PHOTO = 'photo', 'Photograph'
        SELECT = 'select', 'Dropdown'

    form = models.ForeignKey(ExternalRegistrationForm, on_delete=models.CASCADE, related_name='fields')
    field_name = models.CharField(max_length=150)
    field_label = models.CharField(max_length=250)
    field_type = models.CharField(max_length=20, choices=FieldType.choices, default=FieldType.TEXT)
    choices = models.TextField(blank=True, help_text='Comma-separated options for dropdown fields')
    help_text = models.CharField(max_length=300, blank=True)
    is_required = models.BooleanField(default=True)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'pk']

    def __str__(self):
        return self.field_label

    def choice_list(self):
        return [c.strip() for c in self.choices.split(',') if c.strip()]


class ExternalRegistrationSubmission(models.Model):
    form = models.ForeignKey(ExternalRegistrationForm, on_delete=models.CASCADE, related_name='submissions')
    field_data = models.JSONField(default=dict)
    submitted_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-submitted_at']

    def __str__(self):
        name = self.field_data.get('examiner_name_bank') or self.field_data.get('examiner_name') or f'Submission #{self.pk}'
        return f'{name} — {self.form.semester.name}'

    def display_value(self, field_name):
        return self.field_data.get(field_name, '—')


class FacultySubjectAssignment(models.Model):
    """Assign subject faculty per batch (e.g. DVP → A1 PYTHON-II)."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='subject_assignments')
    batch = models.CharField(max_length=20)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='faculty_assignments')
    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='subject_assignments')
    assigned_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='subject_assignments_made')
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['batch', 'subject__name']
        unique_together = ['department', 'batch', 'subject']

    def __str__(self):
        return f'{self.faculty.name} — {self.subject.name} / {self.batch}'


class DutyScheduleUpload(models.Model):
    """Uploaded IPE/GP duty schedule Excel."""
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='duty_uploads')
    schedule_file = models.FileField(upload_to='duty_schedules/')
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    summary = models.JSONField(default=dict, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f'Duty schedule — {self.department.name} ({self.created_at:%d %b %Y})'


class FacultyDutyAssignment(models.Model):
    """Assign faculty (internal or external) to conduct IPE/GP for a subject, batch, and date."""
    class ExamType(models.TextChoices):
        IPE = 'IPE', 'IPE'
        GP = 'GP', 'GP'

    class DutyRole(models.TextChoices):
        INTERNAL = 'INTERNAL', 'Internal'
        EXTERNAL = 'EXTERNAL', 'External'

    faculty = models.ForeignKey(Faculty, on_delete=models.CASCADE, related_name='duty_assignments')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='duty_assignments')
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='duty_assignments')
    exam_type = models.CharField(max_length=5, choices=ExamType.choices)
    duty_role = models.CharField(max_length=10, choices=DutyRole.choices, default=DutyRole.INTERNAL)
    batch = models.CharField(max_length=20, help_text='Batch division e.g. A1')
    duty_date = models.DateField()
    time_slot = models.CharField(max_length=80, blank=True)
    room_no = models.CharField(max_length=50, blank=True)
    schedule_upload = models.ForeignKey(
        DutyScheduleUpload, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='duties',
    )
    assigned_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='duties_assigned')
    is_active = models.BooleanField(default=True)
    marks_locked = models.BooleanField(
        default=False,
        help_text='When set, faculty and department admin cannot edit marks.',
    )
    marks_saved_at = models.DateTimeField(null=True, blank=True)
    marks_locked_at = models.DateTimeField(null=True, blank=True)
    marks_locked_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='duty_marks_locked',
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-duty_date', 'batch']
        unique_together = [
            'faculty', 'department', 'subject', 'exam_type', 'batch', 'duty_date', 'duty_role',
        ]

    def __str__(self):
        return (
            f'{self.faculty.name} ({self.get_duty_role_display()}) — '
            f'{self.exam_type} {self.subject.name} {self.batch} ({self.duty_date})'
        )


class ExamSession(models.Model):
    class ExamType(models.TextChoices):
        IPE = 'IPE', 'Internal Practical Examination'
        GP = 'GP', 'Group Project Evaluation'

    name = models.CharField(max_length=150)
    exam_type = models.CharField(max_length=5, choices=ExamType.choices)
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE, related_name='exam_sessions')
    department = models.ForeignKey(Department, on_delete=models.CASCADE, related_name='exam_sessions')
    date = models.DateField(null=True, blank=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.name} - {self.exam_type}"


class MarkEntry(models.Model):
    student = models.ForeignKey(Student, on_delete=models.CASCADE, related_name='marks')
    exam_session = models.ForeignKey(
        ExamSession, on_delete=models.CASCADE, null=True, blank=True, related_name='marks',
    )
    duty_assignment = models.ForeignKey(
        FacultyDutyAssignment, on_delete=models.CASCADE, null=True, blank=True,
        related_name='mark_entries',
    )
    marks_obtained = models.DecimalField(max_digits=5, decimal_places=2, default=0)
    mark_data = models.JSONField(default=dict, blank=True)
    remarks = models.CharField(max_length=255, blank=True)
    entered_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='marks_entered')
    entered_at = models.DateTimeField(auto_now=True)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['student', 'exam_session'],
                condition=models.Q(exam_session__isnull=False),
                name='unique_mark_per_exam_session',
            ),
            models.UniqueConstraint(
                fields=['student', 'duty_assignment'],
                condition=models.Q(duty_assignment__isnull=False),
                name='unique_mark_per_duty',
            ),
        ]

    def __str__(self):
        return f"{self.student.enrollment_no} - {self.marks_obtained}"


class GeneratedCredential(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='credentials')
    plain_password = models.CharField(max_length=50)
    generated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='credentials_generated')
    generated_at = models.DateTimeField(default=timezone.now)

    def __str__(self):
        return f"{self.user.username}"
