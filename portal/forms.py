from django import forms
from django.contrib.auth.forms import AuthenticationForm
from django.utils import timezone

from .models import (
    User, Semester, Department, Subject, Faculty, Student,
    FormTemplate, GPGroup, ExamSession, MarkEntry, FormField, ProjectCase,
    MarksheetTemplate,
    AttendanceSheetTemplate,
    FacultyDutyAssignment,
    SubjectSyllabus, SubjectPaper,
    ExternalRegistrationForm, ExternalRegistrationField,
)


class LoginForm(AuthenticationForm):
    username = forms.CharField(
        label='Username',
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter username',
        })
    )
    password = forms.CharField(
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Enter password',
        })
    )


class ChangePasswordForm(forms.Form):
    username = forms.CharField(
        label='Username / Enrollment No',
        widget=forms.TextInput(attrs={
            'class': 'form-control form-control-lg',
            'placeholder': 'Enter your username',
            'autocomplete': 'username',
        }),
    )
    current_password = forms.CharField(
        label='Current Password',
        widget=forms.PasswordInput(attrs={
            'class': 'form-control form-control-lg',
            'placeholder': 'Enter current password',
            'autocomplete': 'current-password',
        }),
    )
    new_password = forms.CharField(
        label='New Password',
        min_length=4,
        widget=forms.PasswordInput(attrs={
            'class': 'form-control form-control-lg',
            'placeholder': 'Enter new password',
            'autocomplete': 'new-password',
        }),
    )
    confirm_password = forms.CharField(
        label='Confirm New Password',
        widget=forms.PasswordInput(attrs={
            'class': 'form-control form-control-lg',
            'placeholder': 'Re-enter new password',
            'autocomplete': 'new-password',
        }),
    )

    def clean(self):
        cleaned = super().clean()
        new = cleaned.get('new_password')
        confirm = cleaned.get('confirm_password')
        if new and confirm and new != confirm:
            self.add_error('confirm_password', 'New passwords do not match.')
        return cleaned


class SemesterForm(forms.ModelForm):
    class Meta:
        model = Semester
        fields = ['name', 'academic_year', 'is_active']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'academic_year': forms.TextInput(attrs={'class': 'form-control', 'placeholder': '2025-26'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


class DepartmentForm(forms.ModelForm):
    class Meta:
        model = Department
        fields = ['name', 'code', 'sheet_semester_label', 'sheet_department_label']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'code': forms.TextInput(attrs={'class': 'form-control'}),
            'sheet_semester_label': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'III- ODD-2025'}),
            'sheet_department_label': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'SY- CE\\IT- 1'}),
        }


class AdminUserForm(forms.Form):
    username = forms.CharField(widget=forms.TextInput(attrs={'class': 'form-control'}))
    password = forms.CharField(widget=forms.PasswordInput(attrs={'class': 'form-control'}))
    email = forms.EmailField(required=False, widget=forms.EmailInput(attrs={'class': 'form-control'}))


class SubjectForm(forms.ModelForm):
    class Meta:
        model = Subject
        fields = ['name', 'code', 'subject_type', 'max_marks_ipe', 'max_marks_gp']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'code': forms.TextInput(attrs={'class': 'form-control'}),
            'subject_type': forms.Select(attrs={'class': 'form-select'}),
            'max_marks_ipe': forms.NumberInput(attrs={'class': 'form-control'}),
            'max_marks_gp': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class FacultyForm(forms.ModelForm):
    class Meta:
        model = Faculty
        fields = ['name', 'mentor_code', 'email', 'phone']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'mentor_code': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'phone': forms.TextInput(attrs={'class': 'form-control'}),
        }


class StudentForm(forms.ModelForm):
    class Meta:
        model = Student
        fields = [
            'roll_no', 'name', 'enrollment_no', 'branch', 'batch',
            'mentor', 'student_phone', 'parent_contact',
        ]
        widgets = {
            'roll_no': forms.TextInput(attrs={'class': 'form-control'}),
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'enrollment_no': forms.TextInput(attrs={'class': 'form-control'}),
            'branch': forms.TextInput(attrs={'class': 'form-control'}),
            'batch': forms.TextInput(attrs={'class': 'form-control'}),
            'mentor': forms.TextInput(attrs={'class': 'form-control'}),
            'student_phone': forms.TextInput(attrs={'class': 'form-control'}),
            'parent_contact': forms.TextInput(attrs={'class': 'form-control'}),
        }


class ExcelUploadForm(forms.Form):
    file = forms.FileField(widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls,.csv'}))


class AttendanceTemplateUploadForm(forms.ModelForm):
    class Meta:
        model = AttendanceSheetTemplate
        fields = ['name', 'exam_type', 'template_file']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'exam_type': forms.Select(attrs={'class': 'form-select'}),
            'template_file': forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'}),
        }


class AttendanceDownloadForm(forms.Form):
    exam_type = forms.ChoiceField(
        choices=[('IPE', 'IPE'), ('GP', 'GP')],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_exam_type'}),
    )
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        required=False,
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_att_subject'}),
    )
    subject_selection = forms.ChoiceField(
        required=False,
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_att_subject_selection'}),
        label='GP Subject (Practical combined)',
    )
    batch = forms.ChoiceField(
        required=False,
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_batch'}),
        label='Batch (GP only)',
    )
    semester_label = forms.CharField(
        label='Semester Label',
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'III- ODD-2025'}),
    )
    department_label = forms.CharField(
        label='Department Label',
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'SY- CE\\IT- 1'}),
    )

    def clean(self):
        cleaned = super().clean()
        exam_type = cleaned.get('exam_type')
        if exam_type == 'IPE' and not cleaned.get('subject'):
            self.add_error('subject', 'Select a subject for IPE.')
        if exam_type == 'GP' and not cleaned.get('subject_selection'):
            self.add_error('subject_selection', 'Select a GP subject.')
        return cleaned


class MarksheetTemplateUploadForm(forms.ModelForm):
    replace_existing = forms.BooleanField(
        required=False,
        initial=True,
        label='Replace existing template for this subject',
        widget=forms.CheckboxInput(attrs={'class': 'form-check-input'}),
    )

    class Meta:
        model = MarksheetTemplate
        fields = ['name', 'exam_type', 'subject', 'template_file']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. PYTHON-II IPE Marksheet'}),
            'exam_type': forms.Select(attrs={'class': 'form-select', 'id': 'id_ms_exam_type'}),
            'subject': forms.Select(attrs={'class': 'form-select', 'id': 'id_ms_subject'}),
            'template_file': forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'}),
        }


class MarksheetDownloadForm(forms.Form):
    exam_type = forms.ChoiceField(
        choices=[('IPE', 'IPE'), ('GP', 'GP')],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_ms_dl_exam_type'}),
    )
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        required=False,
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_ms_dl_subject'}),
    )
    subject_selection = forms.ChoiceField(
        required=False,
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_ms_dl_subject_selection'}),
        label='GP Subject (Practical combined)',
    )
    semester_label = forms.CharField(
        label='Semester Label',
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'III- ODD-2025'}),
    )
    department_label = forms.CharField(
        label='Department Label',
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'SY- CE\\IT- 1'}),
    )

    def clean(self):
        cleaned = super().clean()
        exam_type = cleaned.get('exam_type')
        if exam_type == 'IPE' and not cleaned.get('subject'):
            self.add_error('subject', 'Select a subject for IPE.')
        if exam_type == 'GP' and not cleaned.get('subject_selection'):
            self.add_error('subject_selection', 'Select a GP subject.')
        return cleaned


class FacultyDutyAssignmentForm(forms.Form):
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        widget=forms.Select(attrs={'class': 'form-select'}),
    )
    internal_faculty = forms.ChoiceField(
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_ipe_internal_faculty'}),
        label='Internal Faculty',
    )
    internal_faculty_other = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={
            'class': 'form-control text-uppercase',
            'id': 'id_ipe_internal_other',
            'placeholder': 'Enter internal faculty name in CAPITAL letters',
        }),
        label='Other Internal Faculty Name',
    )
    external_faculty_name = forms.CharField(
        widget=forms.TextInput(attrs={
            'class': 'form-control text-uppercase',
            'placeholder': 'Enter external faculty name',
        }),
        label='External Faculty Name',
    )
    batch = forms.ChoiceField(choices=[], widget=forms.Select(attrs={'class': 'form-select'}))
    duty_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}))
    time_slot = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. 9:00 am to 10:30 am'}),
    )
    room_no = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. 410-B'}),
        label='Room No',
    )

    def clean_internal_faculty_other(self):
        return (self.cleaned_data.get('internal_faculty_other') or '').strip().upper()

    def clean_external_faculty_name(self):
        return (self.cleaned_data.get('external_faculty_name') or '').strip().upper()

    def clean(self):
        cleaned = super().clean()
        internal = cleaned.get('internal_faculty')
        other = cleaned.get('internal_faculty_other')
        if internal == 'OTHER' and not other:
            self.add_error('internal_faculty_other', 'Enter the internal faculty name.')
        elif not internal:
            self.add_error('internal_faculty', 'Select internal faculty.')
        return cleaned


class GPDutyAssignmentForm(forms.Form):
    subject_selection = forms.ChoiceField(
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_gp_duty_subject'}),
        label='GP Subject',
    )
    split_choice = forms.ChoiceField(
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_gp_duty_split'}),
        label='Group Split',
    )
    internal_faculty = forms.ChoiceField(
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_gp_internal_faculty'}),
        label='Internal Faculty',
    )
    internal_faculty_other = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={
            'class': 'form-control text-uppercase',
            'id': 'id_gp_internal_other',
            'placeholder': 'Enter faculty name in CAPITAL letters',
            'style': 'display:none;',
        }),
        label='Other Internal Faculty Name',
    )
    external_faculty_name = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'External examiner name',
        }),
        label='External Faculty Name',
    )
    duty_date = forms.DateField(
        widget=forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
    )
    room_no = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. 410-B'}),
        label='Room No',
    )

    def clean_internal_faculty_other(self):
        val = (self.cleaned_data.get('internal_faculty_other') or '').strip().upper()
        return val

    def clean_external_faculty_name(self):
        return (self.cleaned_data.get('external_faculty_name') or '').strip()

    def clean(self):
        cleaned = super().clean()
        internal = cleaned.get('internal_faculty')
        other = cleaned.get('internal_faculty_other', '')
        if internal == 'OTHER' and not other:
            self.add_error('internal_faculty_other', 'Enter the internal faculty name.')
        if internal and internal != 'OTHER' and not other:
            pass
        elif internal != 'OTHER' and not internal:
            self.add_error('internal_faculty', 'Select internal faculty.')
        return cleaned


class DutyScheduleUploadForm(forms.Form):
    schedule_file = forms.FileField(
        widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'}),
        label='Duty Schedule Excel',
    )
    replace_existing = forms.BooleanField(
        required=False, initial=False,
        widget=forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        label='Replace previous Excel-imported duties for this department',
    )

    def clean_schedule_file(self):
        f = self.cleaned_data.get('schedule_file')
        if f and not f.name.lower().endswith(('.xlsx', '.xls')):
            raise forms.ValidationError('Upload an Excel file (.xlsx or .xls).')
        return f


class FinalMarksheetDownloadForm(forms.Form):
    exam_type = forms.ChoiceField(
        choices=[('IPE', 'IPE'), ('GP', 'GP')],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_final_ms_exam_type'}),
    )
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        required=False,
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_final_ms_subject'}),
    )
    subject_selection = forms.ChoiceField(
        required=False,
        choices=[],
        widget=forms.Select(attrs={'class': 'form-select', 'id': 'id_final_ms_subject_selection'}),
        label='GP Subject',
    )
    semester_label = forms.CharField(
        label='Semester Label',
        widget=forms.TextInput(attrs={'class': 'form-control'}),
    )
    department_label = forms.CharField(
        label='Department Label',
        widget=forms.TextInput(attrs={'class': 'form-control'}),
    )
    download_type = forms.ChoiceField(
        choices=[
            ('batch', 'All Batches (separate sheets)'),
            ('combined', 'Combined (single sheet — all batches)'),
        ],
        widget=forms.Select(attrs={'class': 'form-select'}),
        label='Download Format',
        required=False,
        initial='batch',
    )

    def clean(self):
        cleaned = super().clean()
        exam_type = cleaned.get('exam_type')
        if exam_type == 'IPE' and not cleaned.get('subject'):
            self.add_error('subject', 'Select a subject for IPE.')
        if exam_type == 'GP' and not cleaned.get('subject_selection'):
            self.add_error('subject_selection', 'Select a GP subject.')
        return cleaned


FilledMarksheetDownloadForm = FinalMarksheetDownloadForm


SYLLABUS_EXTENSIONS = ('.pdf', '.doc', '.docx')
PAPER_EXTENSIONS = ('.pdf', '.doc', '.docx')


def _validate_extension(file, allowed):
    name = (file.name or '').lower()
    if not any(name.endswith(ext) for ext in allowed):
        raise forms.ValidationError(
            f'Allowed file types: {", ".join(ext.upper().lstrip(".") for ext in allowed)}'
        )


class SyllabusUploadForm(forms.ModelForm):
    replace_existing = forms.BooleanField(
        required=False, initial=False,
        widget=forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        label='Replace existing syllabus for this subject',
    )

    class Meta:
        model = SubjectSyllabus
        fields = ['subject', 'title', 'file']
        widgets = {
            'subject': forms.Select(attrs={'class': 'form-select'}),
            'title': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Optional display title'}),
            'file': forms.FileInput(attrs={'class': 'form-control', 'accept': '.pdf,.doc,.docx'}),
        }

    def clean_file(self):
        f = self.cleaned_data.get('file')
        if f:
            _validate_extension(f, SYLLABUS_EXTENSIONS)
        return f


class PaperUploadForm(forms.ModelForm):
    replace_existing = forms.BooleanField(
        required=False, initial=False,
        widget=forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        label='Replace existing paper for this subject & phase',
    )

    class Meta:
        model = SubjectPaper
        fields = ['subject', 'phase', 'title', 'file']
        widgets = {
            'subject': forms.Select(attrs={'class': 'form-select'}),
            'phase': forms.NumberInput(attrs={'class': 'form-control', 'min': 1, 'max': 20}),
            'title': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. Phase 1 — Programs 1 & 2'}),
            'file': forms.FileInput(attrs={'class': 'form-control', 'accept': '.pdf,.doc,.docx'}),
        }

    def clean_file(self):
        f = self.cleaned_data.get('file')
        if f:
            _validate_extension(f, PAPER_EXTENSIONS)
        return f


class FormTemplateForm(forms.ModelForm):
    class Meta:
        model = FormTemplate
        fields = ['name', 'exam_type']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'exam_type': forms.Select(attrs={'class': 'form-select'}),
        }


class FormFieldForm(forms.ModelForm):
    class Meta:
        model = FormField
        fields = ['field_label', 'field_scope', 'field_type', 'choices', 'is_required', 'order']
        widgets = {
            'field_label': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. Project Guide Name'}),
            'field_scope': forms.Select(attrs={'class': 'form-select'}),
            'field_type': forms.Select(attrs={'class': 'form-select'}),
            'choices': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Option 1, Option 2 (for dropdown only)'}),
            'is_required': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'order': forms.NumberInput(attrs={'class': 'form-control', 'min': 0}),
        }


class ProjectCaseForm(forms.ModelForm):
    class Meta:
        model = ProjectCase
        fields = ['name', 'is_active']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. Case 1'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


class GPDeadlineForm(forms.Form):
    gp_submission_deadline = forms.DateTimeField(
        required=False,
        widget=forms.DateTimeInput(
            attrs={'type': 'datetime-local', 'class': 'form-control'},
            format='%Y-%m-%dT%H:%M',
        ),
        input_formats=['%Y-%m-%dT%H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'],
        help_text='Students cannot submit or edit GP projects after this date and time.',
    )

    def __init__(self, *args, department=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.department = department
        if department and department.gp_submission_deadline:
            local = timezone.localtime(department.gp_submission_deadline)
            self.fields['gp_submission_deadline'].initial = local.strftime('%Y-%m-%dT%H:%M')


class GPGroupForm(forms.ModelForm):
    members = forms.ModelMultipleChoiceField(
        queryset=Student.objects.none(),
        widget=forms.CheckboxSelectMultiple(attrs={'class': 'form-check-input'}),
        required=False,
        label='Group Members',
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['subject'].empty_label = '— Select Subject —'
        self.fields['subject'].required = True
        self.fields['name'].required = True

    class Meta:
        model = GPGroup
        fields = ['subject', 'name', 'members']
        labels = {
            'subject': 'Subject',
            'name': 'Project Title',
        }
        widgets = {
            'subject': forms.Select(attrs={'class': 'form-select'}),
            'name': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Enter project title for the selected subject',
            }),
        }


class ExamSessionForm(forms.ModelForm):
    class Meta:
        model = ExamSession
        fields = ['name', 'exam_type', 'subject', 'date', 'is_active']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'exam_type': forms.Select(attrs={'class': 'form-select'}),
            'subject': forms.Select(attrs={'class': 'form-select'}),
            'date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


class MarkEntryForm(forms.ModelForm):
    class Meta:
        model = MarkEntry
        fields = ['marks_obtained', 'remarks']
        widgets = {
            'marks_obtained': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.5'}),
            'remarks': forms.TextInput(attrs={'class': 'form-control'}),
        }


def _widget_for_field(field):
    attrs = {'class': 'form-control'}
    if field.field_type == FormField.FieldType.TEXTAREA:
        return forms.Textarea(attrs={**attrs, 'rows': 3})
    if field.field_type == FormField.FieldType.NUMBER:
        return forms.NumberInput(attrs=attrs)
    if field.field_type == FormField.FieldType.YES_NO:
        return forms.Select(attrs={'class': 'form-select'}, choices=[('', '— Select —'), ('Yes', 'Yes'), ('No', 'No')])
    if field.field_type == FormField.FieldType.SELECT:
        choices = [('', '— Select —')] + [(c, c) for c in field.choice_list()]
        return forms.Select(attrs={'class': 'form-select'}, choices=choices)
    return forms.TextInput(attrs=attrs)


def build_dynamic_form(template, initial=None, scope=None):
    """Build a dynamic form from FormTemplate fields."""
    fields_dict = {}
    qs = template.fields.all()
    if scope:
        qs = qs.filter(field_scope=scope)
    for field in qs:
        widget = _widget_for_field(field)
        if field.field_type == FormField.FieldType.NUMBER:
            fields_dict[field.field_name] = forms.DecimalField(
                label=field.field_label, required=field.is_required, widget=widget
            )
        else:
            fields_dict[field.field_name] = forms.CharField(
                label=field.field_label, required=field.is_required, widget=widget
            )
    DynamicForm = type('DynamicForm', (forms.Form,), fields_dict)
    return DynamicForm(initial=initial or {})


class ExternalRegistrationFormCreateForm(forms.ModelForm):
    class Meta:
        model = ExternalRegistrationForm
        fields = ['title', 'semester', 'description', 'is_active']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control'}),
            'semester': forms.Select(attrs={'class': 'form-select'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


class ExternalRegistrationFieldForm(forms.ModelForm):
    class Meta:
        model = ExternalRegistrationField
        fields = ['field_label', 'field_type', 'choices', 'help_text', 'is_required', 'order']
        widgets = {
            'field_label': forms.TextInput(attrs={'class': 'form-control'}),
            'field_type': forms.Select(attrs={'class': 'form-select'}),
            'choices': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Regular,Jain'}),
            'help_text': forms.TextInput(attrs={'class': 'form-control'}),
            'is_required': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
        }


class IPEInvitationCreateForm(forms.Form):
    excel_file = forms.FileField(
        label='Excel — External Faculty List',
        widget=forms.ClearableFileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'}),
        help_text='Columns: Name, Designation, College Name, City/State, Email',
    )
    letter_date = forms.CharField(
        label='Date',
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'e.g. 17 Feb. 2025',
        }),
    )
    subject_line = forms.CharField(
        label='Sub (Subject line)',
        initial='Appointment as an External Examiner for conducting Internal Practical Examination of B.E. Semester-III',
        widget=forms.TextInput(attrs={'class': 'form-control'}),
    )
    practical_date = forms.CharField(
        label='Date of Practical',
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'e.g. 20 February 2025',
        }),
    )
    branch = forms.CharField(
        label='Branch',
        initial='CE/IT/CSD/AIML/AIDS/RAI/CSE/CST/CSIT/CEA',
        widget=forms.TextInput(attrs={'class': 'form-control'}),
    )
    exam_time = forms.CharField(
        label='Exam Time',
        initial='9:00 am onwards',
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'e.g. 9:00 am onwards',
        }),
    )
