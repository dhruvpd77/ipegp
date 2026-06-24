"""External examiner registration form — field seeding and dynamic form builder."""
from django import forms
from django.core.files.uploadedfile import UploadedFile

from .models import ExternalRegistrationField, ExternalRegistrationForm
from .gp_utils import slugify_field_name

DEFAULT_EXTERNAL_FIELDS = [
    ('SR.NO', 'sr_no', 'number', '', True, 1),
    ('Email Address', 'email_address', 'email', '', True, 2),
    ('Examiner Name As per Bank Account', 'examiner_name_bank', 'text', '', True, 3),
    (
        'Date (Available on both or either of the dates)',
        'available_date', 'date',
        'Select a date you are available for examination',
        True, 4,
    ),
    ('Your recent Photograph', 'photograph', 'photo', 'Upload a recent passport-size photograph', True, 5),
    ('Qualification', 'qualification', 'text', '', True, 6),
    ('Organization', 'organization', 'text', '', True, 7),
    ('Designation', 'designation', 'text', '', True, 8),
    ('Experience', 'experience', 'textarea', '', True, 9),
    ('Email', 'email', 'email', '', True, 10),
    ('Contact No', 'contact_no', 'text', '', True, 11),
    ('Bank Account Number', 'bank_account_number', 'text', '', True, 12),
    ('Bank Name', 'bank_name', 'text', '', True, 13),
    ('IFSC Code', 'ifsc_code', 'text', '', True, 14),
    ('Bank Branch', 'bank_branch', 'text', '', True, 15),
    ('Address', 'address', 'textarea', '', True, 16),
    (
        'TA in KM (One Way)',
        'ta_km_one_way', 'number',
        'Travel allowance distance in KM (one way)',
        False, 17,
    ),
    (
        'Amount for TA (1KM = Rs 8)',
        'ta_amount', 'number',
        'Auto-calculated as KM × 8 if left blank',
        False, 18,
    ),
    ('Food (Reg/Jain)', 'food_preference', 'select', '', True, 19),
]


def seed_external_form_fields(reg_form):
    """Create default external examiner fields if form has none."""
    if reg_form.fields.exists():
        return 0
    created = 0
    for label, name, ftype, help_text, required, order in DEFAULT_EXTERNAL_FIELDS:
        choices = 'Regular,Jain' if name == 'food_preference' else ''
        ExternalRegistrationField.objects.create(
            form=reg_form,
            field_name=name,
            field_label=label,
            field_type=ftype,
            choices=choices,
            help_text=help_text,
            is_required=required,
            order=order,
        )
        created += 1
    return created


def _widget_for_external_field(field):
    attrs = {'class': 'form-control ext-input'}
    if field.field_type == ExternalRegistrationField.FieldType.TEXTAREA:
        return forms.Textarea(attrs={**attrs, 'rows': 3})
    if field.field_type == ExternalRegistrationField.FieldType.NUMBER:
        return forms.NumberInput(attrs=attrs)
    if field.field_type == ExternalRegistrationField.FieldType.DATE:
        return forms.DateInput(attrs={**attrs, 'type': 'date'})
    if field.field_type == ExternalRegistrationField.FieldType.EMAIL:
        return forms.EmailInput(attrs=attrs)
    if field.field_type == ExternalRegistrationField.FieldType.SELECT:
        choices = [('', '— Select —')] + [(c, c) for c in field.choice_list()]
        return forms.Select(attrs={'class': 'form-select ext-input'}, choices=choices)
    if field.field_type == ExternalRegistrationField.FieldType.PHOTO:
        return forms.ClearableFileInput(attrs={'class': 'form-control ext-input', 'accept': 'image/*'})
    return forms.TextInput(attrs=attrs)


def build_external_registration_form(reg_form):
    """Build a Django Form class from ExternalRegistrationField definitions."""
    fields_dict = {}
    for field in reg_form.fields.all():
        widget = _widget_for_external_field(field)
        label = field.field_label
        help_text = field.help_text or None
        required = field.is_required

        if field.field_type == ExternalRegistrationField.FieldType.PHOTO:
            fields_dict[field.field_name] = forms.ImageField(
                label=label, required=required, widget=widget, help_text=help_text,
            )
        elif field.field_type == ExternalRegistrationField.FieldType.NUMBER:
            fields_dict[field.field_name] = forms.DecimalField(
                label=label, required=required, widget=widget, help_text=help_text,
                min_value=0, decimal_places=2,
            )
        elif field.field_type == ExternalRegistrationField.FieldType.EMAIL:
            fields_dict[field.field_name] = forms.EmailField(
                label=label, required=required, widget=widget, help_text=help_text,
            )
        elif field.field_type == ExternalRegistrationField.FieldType.DATE:
            fields_dict[field.field_name] = forms.DateField(
                label=label, required=required, widget=widget, help_text=help_text,
            )
        else:
            fields_dict[field.field_name] = forms.CharField(
                label=label, required=required, widget=widget, help_text=help_text,
            )
    return type('ExternalRegistrationDynamicForm', (forms.Form,), fields_dict)


def save_external_submission(reg_form, cleaned_data, files):
    """Persist submission; store photo paths in field_data."""
    from django.core.files.storage import default_storage

    from .models import ExternalRegistrationSubmission

    data = {}
    photo_uploads = {}
    for field in reg_form.fields.all():
        key = field.field_name
        if field.field_type == ExternalRegistrationField.FieldType.PHOTO:
            upload = files.get(key) or cleaned_data.get(key)
            if upload and isinstance(upload, UploadedFile):
                photo_uploads[key] = upload
        else:
            val = cleaned_data.get(key)
            if val is not None:
                data[key] = str(val) if val != '' else ''

    if not data.get('ta_amount') and data.get('ta_km_one_way'):
        try:
            km = float(data['ta_km_one_way'])
            data['ta_amount'] = str(round(km * 8, 2))
        except (TypeError, ValueError):
            pass

    sub = ExternalRegistrationSubmission.objects.create(form=reg_form, field_data=data)
    for key, upload in photo_uploads.items():
        ext = upload.name.rsplit('.', 1)[-1] if '.' in upload.name else 'jpg'
        path = f'external_submissions/{sub.pk}_{key}.{ext}'
        data[key] = default_storage.save(path, upload)
    sub.field_data = data
    sub.save(update_fields=['field_data'])
    return sub


def unique_field_name(form, label):
    base = slugify_field_name(label)
    name = base
    n = 1
    while form.fields.filter(field_name=name).exists():
        name = f'{base}_{n}'
        n += 1
    return name
