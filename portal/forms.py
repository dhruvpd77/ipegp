from django import forms
from django.contrib.auth.forms import AuthenticationForm

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
        widget=forms.Select(attrs={'class': 'form-select'}),
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
        widget=forms.Select(attrs={'class': 'form-select'}),
    )
    semester_label = forms.CharField(
        label='Semester Label',
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'III- ODD-2025'}),
    )
    department_label = forms.CharField(
        label='Department Label',
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'SY- CE\\IT- 1'}),
    )


class FacultyDutyAssignmentForm(forms.ModelForm):
    batch = forms.ChoiceField(choices=[], widget=forms.Select(attrs={'class': 'form-select'}))

    class Meta:
        model = FacultyDutyAssignment
        fields = ['exam_type', 'subject', 'faculty', 'duty_role', 'batch', 'duty_date', 'time_slot', 'room_no']
        widgets = {
            'exam_type': forms.Select(attrs={'class': 'form-select', 'id': 'id_duty_exam_type'}),
            'subject': forms.Select(attrs={'class': 'form-select'}),
            'faculty': forms.Select(attrs={'class': 'form-select'}),
            'duty_role': forms.Select(attrs={'class': 'form-select'}),
            'duty_date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'time_slot': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. 9:00 am to 10:30 am'}),
            'room_no': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. 410-B'}),
        }


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
        widget=forms.Select(attrs={'class': 'form-select'}),
    )
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        widget=forms.Select(attrs={'class': 'form-select'}),
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
            'name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'e.g. Case A - Web Application'}),
            'is_active': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


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
