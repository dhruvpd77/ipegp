import json

from django.contrib import messages
from django.contrib.auth import authenticate, login, logout, update_session_auth_hash
from django.contrib.auth.decorators import login_required
from django.db import IntegrityError
from django.db.models import Q
import os
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.utils import timezone
from django.http import HttpResponse, FileResponse, HttpResponseForbidden
from django.views.decorators.clickjacking import xframe_options_sameorigin
import mimetypes

from .decorators import role_required
from .forms import (
    LoginForm, ChangePasswordForm, SemesterForm, DepartmentForm, AdminUserForm, SubjectForm,
    FacultyForm, StudentForm, ExcelUploadForm, FormTemplateForm, GPGroupForm,
    FormFieldForm, ProjectCaseForm, GPDeadlineForm,
    ExamSessionForm, MarkEntryForm, build_dynamic_form,
    AttendanceTemplateUploadForm, AttendanceDownloadForm,
    MarksheetTemplateUploadForm, MarksheetDownloadForm,
    FacultyDutyAssignmentForm, FinalMarksheetDownloadForm,
    SyllabusUploadForm, PaperUploadForm, DutyScheduleUploadForm,
    ExternalRegistrationFormCreateForm, ExternalRegistrationFieldForm,
)
from .external_form_utils import (
    build_external_registration_form, save_external_submission,
    seed_external_form_fields, unique_field_name,
)
from .gp_utils import (
    get_gp_template, get_project_cases, get_group_fields, get_member_fields,
    save_gp_submission, get_faculty_map_for_subjects, get_taken_members_by_subject,
    export_gp_submissions_excel, create_default_gp_template,
    import_fields_from_excel, import_cases_from_excel,
    build_gp_subject_selection_options, get_bundle_groups,
    bundle_subject_entries_from_groups, selection_value_for_groups,
    subjects_for_department, get_taken_titles_by_subject,
    parse_subject_selection,
    get_pending_gp_students,
    _resolve_subject_entry_from_post,
    is_gp_submission_locked,
    GENDER_CHOICES, YES_NO_CHOICES,
    MAX_GP_GROUP_MEMBERS,
)
from .attendance_sheet import (
    generate_attendance_workbook, generate_gp_attendance_workbook,
    get_default_template_path, ensure_default_templates,
)
from .duty_schedule import import_duty_schedule, department_batches
from .marksheet import (
    parse_marksheet_template, resolve_marksheet_template,
    generate_marksheet_workbook, generate_compiled_marksheet_workbook,
    ensure_default_marksheet_template,
    save_duty_marks, build_duty_marksheet_page_context, get_duty_marks_status,
    can_edit_duty_marks, verify_and_lock_duty_marks,
)
from .models import (
    User, Semester, SemesterAdminAssignment, Department, DepartmentAdminAssignment,
    Subject, Faculty, Student, FormTemplate, FormField, GPGroup, GPGroupMemberDetail,
    FormSubmission, ProjectCase,
    ExamSession, MarkEntry, GeneratedCredential, AttendanceSheetTemplate,
    MarksheetTemplate, FacultyDutyAssignment, DutyScheduleUpload,
    SubjectSyllabus, SubjectPaper, FacultySubjectAssignment,
    ExternalRegistrationForm, ExternalRegistrationField, ExternalRegistrationSubmission,
)
from .utils import (
    import_students_from_excel, import_faculty_from_excel,
    create_form_fields_from_excel, export_credentials_excel,
    export_students_excel, sort_students_by_roll,
)
from .semester_access import (
    can_manage_semester_status,
    get_user_semester,
    is_portal_user_blocked_by_inactive_semester,
)


def resolve_semester(user, ctx, request=None):
    """Resolve the active semester for the current user/request."""
    semester = ctx.get('semester')
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        if request is not None:
            semester_id = request.POST.get('semester_id') or request.GET.get('semester')
            if semester_id:
                return get_object_or_404(Semester, pk=semester_id)
        if semester:
            return semester
        return Semester.objects.filter(is_active=True).order_by('-created_at').first()
    return semester


def can_select_department(user):
    return user.is_superuser or user.role in (User.Role.SUPER_ADMIN, User.Role.SEMESTER_ADMIN)


def get_departments_for_user(user, ctx, request=None):
    """Return departments available in dropdowns for the current user."""
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return Department.objects.select_related('semester').order_by('semester__name', 'name')
    if user.role == User.Role.SEMESTER_ADMIN:
        semester = resolve_semester(user, ctx, request)
        if semester:
            return Department.objects.filter(semester=semester).order_by('name')
    dept = ctx.get('department')
    if dept:
        return Department.objects.filter(pk=dept.pk)
    return Department.objects.none()


def _department_in_scope(user, ctx, dept):
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return True
    if user.role == User.Role.SEMESTER_ADMIN:
        semester = ctx.get('semester')
        return semester and dept.semester_id == semester.pk
    assigned = ctx.get('department')
    return assigned and dept.pk == assigned.pk


def resolve_department(user, ctx, request=None):
    """Resolve the active department. None means all departments in scope."""
    dept = ctx.get('department')
    is_super = user.is_superuser or user.role == User.Role.SUPER_ADMIN
    is_semester_admin = user.role == User.Role.SEMESTER_ADMIN

    if user.role == User.Role.DEPARTMENT_ADMIN and dept:
        return dept

    if request is not None:
        if 'department_id' in request.POST or 'department' in request.GET:
            dept_param = request.POST.get('department_id') or request.GET.get('department')
            if dept_param in ('all', '', None):
                return None
            dept = get_object_or_404(Department, pk=dept_param)
            if not _department_in_scope(user, ctx, dept):
                return None
            return dept

    if is_semester_admin:
        return None

    if is_super:
        semester = resolve_semester(user, ctx, request)
        if semester:
            return Department.objects.filter(semester=semester).order_by('name').first()
        return Department.objects.order_by('name').first()

    return dept


def dept_filter_context(user, ctx, request, dept):
    departments = get_departments_for_user(user, ctx, request)
    return {
        'department': dept,
        'departments': departments,
        'show_dept_filter': can_select_department(user) and departments.exists(),
        'is_super': user.is_superuser or user.role == User.Role.SUPER_ADMIN,
    }


def students_in_scope(user, ctx, dept=None):
    if dept:
        return Student.objects.filter(department=dept)
    if user.role == User.Role.SEMESTER_ADMIN:
        semester = ctx.get('semester')
        if semester:
            return Student.objects.filter(department__semester=semester)
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return Student.objects.all()
    assigned = ctx.get('department')
    return Student.objects.filter(department=assigned) if assigned else Student.objects.none()


def faculty_for_department(dept):
    """Internal departmental faculty + faculty with active duty in this department."""
    duty_qs = FacultyDutyAssignment.objects.filter(
        department=dept, is_active=True,
    )
    roster = department_batches(dept)
    if roster:
        duty_qs = duty_qs.filter(batch__in=roster)
    duty_faculty_ids = duty_qs.values_list('faculty_id', flat=True)
    return Faculty.objects.filter(
        Q(department=dept, is_external=False) |
        Q(pk__in=duty_faculty_ids)
    ).distinct()


def faculty_in_scope(user, ctx, dept=None):
    if dept:
        return faculty_for_department(dept)
    if user.role == User.Role.SEMESTER_ADMIN:
        semester = ctx.get('semester')
        if semester:
            return Faculty.objects.filter(department__semester=semester)
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return Faculty.objects.all()
    assigned = ctx.get('department')
    return Faculty.objects.filter(department=assigned) if assigned else Faculty.objects.none()


def submissions_in_scope(user, ctx, dept=None):
    base = FormSubmission.objects.all()
    if dept:
        return base.filter(Q(group__department=dept) | Q(student__department=dept))
    if user.role == User.Role.SEMESTER_ADMIN:
        semester = ctx.get('semester')
        if semester:
            return base.filter(
                Q(group__department__semester=semester) | Q(student__department__semester=semester)
            )
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return base
    assigned = ctx.get('department')
    if assigned:
        return base.filter(Q(group__department=assigned) | Q(student__department=assigned))
    return base.none()


def _dept_query_param(dept):
    return f'department={dept.pk}' if dept else 'department=all'


def get_user_context(user):
    ctx = {'semester': None, 'department': None}
    if user.role == User.Role.SEMESTER_ADMIN:
        try:
            ctx['semester'] = user.semester_admin_assignment.semester
        except SemesterAdminAssignment.DoesNotExist:
            pass
    elif user.role == User.Role.DEPARTMENT_ADMIN:
        try:
            ctx['department'] = user.department_admin_assignment.department
            ctx['semester'] = ctx['department'].semester
        except DepartmentAdminAssignment.DoesNotExist:
            pass
    elif user.role == User.Role.FACULTY:
        try:
            ctx['department'] = user.faculty_profile.department
            ctx['semester'] = ctx['department'].semester
        except Faculty.DoesNotExist:
            pass
    elif user.role == User.Role.STUDENT:
        try:
            ctx['department'] = user.student_profile.department
            ctx['semester'] = ctx['department'].semester
        except Student.DoesNotExist:
            pass
    return ctx


def _post_login_redirect(request):
    if is_portal_user_blocked_by_inactive_semester(request.user):
        return redirect('portal:semester_inactive')
    return redirect('portal:dashboard')


def home(request):
    if request.user.is_authenticated:
        return _post_login_redirect(request)
    return _render_login(request)


def login_view(request, role=None):
    if request.user.is_authenticated:
        return _post_login_redirect(request)
    return _render_login(request)


def _render_login(request):
    if request.method == 'POST':
        form = LoginForm(request, data=request.POST)
        if form.is_valid():
            login(request, form.get_user())
            return _post_login_redirect(request)
    else:
        form = LoginForm()
    return render(request, 'portal/home.html', {'form': form})


def logout_view(request):
    logout(request)
    messages.info(request, 'You have been logged out.')
    return redirect('portal:home')


def change_password(request):
    """Change password with username + current password (login page or while logged in)."""
    if request.method == 'POST':
        form = ChangePasswordForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data['username'].strip()
            user = authenticate(
                request,
                username=username,
                password=form.cleaned_data['current_password'],
            )
            if user is None:
                messages.error(request, 'Invalid username or current password.')
            elif request.user.is_authenticated and request.user.pk != user.pk:
                messages.error(request, 'You can only change your own password.')
            else:
                user.set_password(form.cleaned_data['new_password'])
                user.save()
                if request.user.is_authenticated and request.user.pk == user.pk:
                    update_session_auth_hash(request, user)
                    messages.success(request, 'Password updated successfully.')
                    return redirect('portal:dashboard')
                messages.success(request, 'Password changed successfully. Please sign in with your new password.')
                return redirect('portal:home')
    else:
        initial = {}
        if request.user.is_authenticated:
            initial['username'] = request.user.username
        form = ChangePasswordForm(initial=initial)
    return render(request, 'portal/change_password.html', {
        'form': form,
        'logged_in': request.user.is_authenticated,
    })


@login_required
def dashboard(request):
    user = request.user
    ctx = get_user_context(user)

    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return render(request, 'portal/dashboards/super_admin.html', {
            'semesters': Semester.objects.all(),
            'total_students': Student.objects.count(),
            'total_faculty': Faculty.objects.count(),
            'total_departments': Department.objects.count(),
        })
    elif user.role == User.Role.SEMESTER_ADMIN:
        semester = ctx.get('semester')
        dept = resolve_department(user, ctx, request)
        batch_filter = request.GET.get('batch', '')
        all_students = students_in_scope(user, ctx, dept)
        students = all_students.filter(batch=batch_filter) if batch_filter else all_students
        batches = all_students.values_list('batch', flat=True).distinct().order_by('batch')
        all_depts = Department.objects.filter(semester=semester) if semester else Department.objects.none()
        return render(request, 'portal/dashboards/semester_admin.html', {
            'semester': semester,
            'department': dept,
            'departments': all_depts,
            'show_dept_filter': all_depts.exists(),
            'subjects': Subject.objects.filter(semester=semester) if semester else Subject.objects.none(),
            'student_count': students.count(),
            'faculty_count': faculty_in_scope(user, ctx, dept).count(),
            'submission_count': submissions_in_scope(user, ctx, dept).count(),
            'recent_students': students.select_related('department').order_by('batch', 'roll_no')[:8],
            'batches': batches,
            'batch_filter': batch_filter,
        })
    elif user.role == User.Role.DEPARTMENT_ADMIN:
        dept = ctx.get('department')
        batch_filter = request.GET.get('batch', '')
        all_students = Student.objects.filter(department=dept)
        students = all_students.filter(batch=batch_filter) if batch_filter else all_students
        batches = all_students.values_list('batch', flat=True).distinct().order_by('batch')
        return render(request, 'portal/dashboards/department_admin.html', {
            'department': dept,
            'student_count': students.count(),
            'faculty_count': Faculty.objects.filter(department=dept).count(),
            'submission_count': FormSubmission.objects.filter(
                Q(group__department=dept) | Q(student__department=dept)
            ).count(),
            'recent_students': students.order_by('batch', 'roll_no')[:8],
            'batches': batches,
            'batch_filter': batch_filter,
        })
    elif user.role == User.Role.FACULTY:
        dept = ctx.get('department')
        faculty = user.faculty_profile
        duty_count = FacultyDutyAssignment.objects.filter(
            faculty=faculty, is_active=True,
        ).count()
        exam_sessions = ExamSession.objects.filter(department=dept, is_active=True)
        return render(request, 'portal/dashboards/faculty.html', {
            'faculty': faculty,
            'department': dept,
            'duty_count': duty_count,
            'exam_sessions': exam_sessions,
        })
    elif user.role == User.Role.STUDENT:
        student = user.student_profile
        return render(request, 'portal/dashboards/student.html', {
            'student': student,
        })
    return redirect('portal:home')


# ─── Super Admin Views ───

@login_required
def semester_inactive(request):
    semester = get_user_semester(request.user)
    if not is_portal_user_blocked_by_inactive_semester(request.user):
        return redirect('portal:dashboard')
    return render(request, 'portal/semester_inactive.html', {'semester': semester})


@login_required
def semester_toggle_active(request, pk):
    sem = get_object_or_404(Semester, pk=pk)
    if not can_manage_semester_status(request.user, sem):
        return HttpResponseForbidden('You cannot change this semester status.')
    if request.method == 'POST':
        sem.is_active = not sem.is_active
        sem.save(update_fields=['is_active'])
        state = 'activated' if sem.is_active else 'deactivated'
        messages.success(request, f'Semester "{sem.name}" has been {state}.')
        next_url = request.POST.get('next')
        if next_url:
            return redirect(next_url)
        return redirect('portal:dashboard')
    return redirect('portal:semester_list')


@role_required(User.Role.SUPER_ADMIN)
def semester_list(request):
    semesters = Semester.objects.all()
    form = SemesterForm()
    return render(request, 'portal/super_admin/semesters.html', {'semesters': semesters, 'form': form})


@role_required(User.Role.SUPER_ADMIN)
def semester_create(request):
    if request.method == 'POST':
        form = SemesterForm(request.POST)
        if form.is_valid():
            sem = form.save(commit=False)
            sem.created_by = request.user
            sem.save()
            messages.success(request, f'Semester "{sem.name}" created successfully.')
            return redirect('portal:semester_list')
    return redirect('portal:semester_list')


@role_required(User.Role.SUPER_ADMIN)
def semester_edit(request, pk):
    sem = get_object_or_404(Semester, pk=pk)
    if request.method == 'POST':
        form = SemesterForm(request.POST, instance=sem)
        if form.is_valid():
            form.save()
            messages.success(request, 'Semester updated.')
            return redirect('portal:semester_list')
    form = SemesterForm(instance=sem)
    return render(request, 'portal/super_admin/semester_edit.html', {'form': form, 'semester': sem})


@role_required(User.Role.SUPER_ADMIN)
def semester_delete(request, pk):
    sem = get_object_or_404(Semester, pk=pk)
    if request.method == 'POST':
        sem.delete()
        messages.success(request, 'Semester deleted.')
    return redirect('portal:semester_list')


@role_required(User.Role.SUPER_ADMIN)
def create_semester_admin(request, semester_pk):
    semester = get_object_or_404(Semester, pk=semester_pk)
    if request.method == 'POST':
        form = AdminUserForm(request.POST)
        if form.is_valid():
            user = User.objects.create_user(
                username=form.cleaned_data['username'],
                password=form.cleaned_data['password'],
                email=form.cleaned_data.get('email', ''),
                role=User.Role.SEMESTER_ADMIN,
            )
            SemesterAdminAssignment.objects.create(user=user, semester=semester)
            messages.success(request, f'Semester admin "{user.username}" created.')
            return redirect('portal:semester_list')
    else:
        form = AdminUserForm()
    return render(request, 'portal/super_admin/create_admin.html', {
        'form': form, 'title': f'Create Semester Admin for {semester.name}', 'back_url': 'portal:semester_list'
    })


# ─── Semester Admin Views ───

@role_required(User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def department_list(request):
    ctx = get_user_context(request.user)
    semester = resolve_semester(request.user, ctx, request)
    is_super = request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN

    if request.method == 'POST':
        form = DepartmentForm(request.POST)
        if not semester:
            messages.error(request, 'Please create a semester first before adding departments.')
        elif form.is_valid():
            try:
                dept = form.save(commit=False)
                dept.semester = semester
                dept.created_by = request.user
                dept.save()
                messages.success(request, f'Department "{dept.name}" created successfully.')
                if is_super and semester:
                    return redirect(f'{reverse("portal:department_list")}?semester={semester.pk}')
                return redirect('portal:department_list')
            except IntegrityError:
                messages.error(request, f'Department "{form.cleaned_data["name"]}" already exists for this semester.')
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = DepartmentForm()

    if semester:
        departments = Department.objects.filter(semester=semester)
    elif is_super:
        departments = Department.objects.all()
    else:
        departments = Department.objects.none()

    return render(request, 'portal/semester_admin/departments.html', {
        'departments': departments,
        'form': form,
        'semester': semester,
        'semesters': Semester.objects.all() if is_super else None,
        'is_super': is_super,
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def department_create(request):
    """Legacy POST endpoint — redirects to department_list handler."""
    return department_list(request)


@role_required(User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def department_delete(request, pk):
    dept = get_object_or_404(Department, pk=pk)
    semester_pk = dept.semester_id
    if request.method == 'POST':
        dept.delete()
        messages.success(request, 'Department deleted.')
    is_super = request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN
    if is_super:
        return redirect(f'{reverse("portal:department_list")}?semester={semester_pk}')
    return redirect('portal:department_list')


@role_required(User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def create_department_admin(request, dept_pk):
    dept = get_object_or_404(Department, pk=dept_pk)
    if request.method == 'POST':
        form = AdminUserForm(request.POST)
        if form.is_valid():
            user = User.objects.create_user(
                username=form.cleaned_data['username'],
                password=form.cleaned_data['password'],
                email=form.cleaned_data.get('email', ''),
                role=User.Role.DEPARTMENT_ADMIN,
            )
            DepartmentAdminAssignment.objects.create(user=user, department=dept)
            messages.success(request, f'Department admin "{user.username}" created for {dept.name}.')
            return redirect('portal:department_list')
    else:
        form = AdminUserForm()
    return render(request, 'portal/super_admin/create_admin.html', {
        'form': form, 'title': f'Create Department Admin for {dept.name}', 'back_url': 'portal:department_list'
    })


# ─── User Management Hub ───

@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def user_management(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    context = dept_filter_context(request.user, ctx, request, dept)
    return render(request, 'portal/admin/user_management.html', context)


# ─── Subject Management ───

@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def subject_list(request):
    ctx = get_user_context(request.user)
    is_super = request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN
    is_dept_admin = request.user.role == User.Role.DEPARTMENT_ADMIN

    if is_dept_admin:
        dept = ctx['department']
        semester = dept.semester
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
    else:
        semester = resolve_semester(request.user, ctx, request)
        subjects = Subject.objects.filter(semester=semester) if semester else Subject.objects.all()

    if request.method == 'POST':
        form = SubjectForm(request.POST)
        if form.is_valid():
            subj = form.save(commit=False)
            if is_dept_admin:
                subj.department = ctx['department']
                subj.semester = ctx['department'].semester
            else:
                subj.semester = resolve_semester(request.user, ctx, request)
                if not subj.semester:
                    messages.error(request, 'Please select a semester before adding a subject.')
                    return render(request, 'portal/subjects/list.html', {
                        'subjects': subjects, 'form': form, 'semester': semester,
                        'semesters': Semester.objects.all() if is_super else None,
                        'is_super': is_super,
                    })
            subj.save()
            messages.success(request, f'Subject "{subj.name}" added successfully.')
            if is_super and subj.semester_id:
                return redirect(f'{reverse("portal:subject_list")}?semester={subj.semester_id}')
            return redirect('portal:subject_list')
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = SubjectForm()

    return render(request, 'portal/subjects/list.html', {
        'subjects': subjects,
        'form': form,
        'semester': semester,
        'semesters': Semester.objects.all() if is_super else None,
        'is_super': is_super,
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def subject_create(request):
    """Legacy POST endpoint — redirects to subject_list handler."""
    return subject_list(request)


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def subject_edit(request, pk):
    subj = get_object_or_404(Subject, pk=pk)
    if request.method == 'POST':
        form = SubjectForm(request.POST, instance=subj)
        if form.is_valid():
            form.save()
            messages.success(request, 'Subject updated.')
            return redirect('portal:subject_list')
    form = SubjectForm(instance=subj)
    return render(request, 'portal/subjects/edit.html', {'form': form, 'subject': subj})


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def subject_delete(request, pk):
    subj = get_object_or_404(Subject, pk=pk)
    if request.method == 'POST':
        subj.delete()
        messages.success(request, 'Subject deleted.')
    return redirect('portal:subject_list')


# ─── Department Admin: Students & Faculty ───

@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def upload_students(request):
    """Legacy URL — redirect to Students page with upload section."""
    dept_id = request.GET.get('department') or request.POST.get('department_id')
    url = reverse('portal:student_list')
    if dept_id:
        url += f'?department={dept_id}#upload-students'
    else:
        url += '#upload-students'
    return redirect(url)


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def upload_faculty(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    departments = get_departments_for_user(request.user, ctx, request)

    if request.method == 'POST' and dept:
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            created, updated = import_faculty_from_excel(form.cleaned_data['file'], dept)
            messages.success(request, f'Imported {created} new faculty, updated {updated}.')
            return redirect('portal:faculty_list')
    else:
        form = ExcelUploadForm()
    return render(request, 'portal/department_admin/upload_faculty.html', {
        'form': form,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def faculty_list(request):
    ctx = get_user_context(request.user)
    is_super = request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN
    dept = resolve_department(request.user, ctx, request)
    departments = get_departments_for_user(request.user, ctx, request)
    search = request.GET.get('q', '').strip()

    if request.method == 'POST' and request.POST.get('form_type') == 'add_faculty':
        form = FacultyForm(request.POST)
        target_dept = resolve_department(request.user, ctx, request)
        if not target_dept:
            messages.error(request, 'Please select a department first.')
        elif form.is_valid():
            fac = form.save(commit=False)
            fac.department = target_dept
            fac.save()
            messages.success(request, f'Faculty "{fac.name}" added successfully.')
            return redirect(_faculty_redirect(request, target_dept))
        else:
            messages.error(request, 'Please correct the errors below.')
    else:
        form = FacultyForm()

    faculty = faculty_in_scope(request.user, ctx, dept).select_related('department')
    if search:
        faculty = faculty.filter(
            Q(name__icontains=search) | Q(mentor_code__icontains=search) |
            Q(email__icontains=search) | Q(phone__icontains=search)
        )

    return render(request, 'portal/faculty/list.html', {
        'faculty_list': faculty,
        'form': form,
        'search': search,
        **dept_filter_context(request.user, ctx, request, dept),
    })


def _faculty_redirect(request, dept):
    url = reverse('portal:faculty_list')
    params = [_dept_query_param(dept)] if can_select_department(request.user) else []
    if dept and not can_select_department(request.user):
        params = [f'department={dept.pk}']
    if request.GET.get('q'):
        params.append(f'q={request.GET.get("q")}')
    return f'{url}?{"&".join(params)}' if params else url


def _student_redirect(request, dept):
    url = reverse('portal:student_list')
    params = [_dept_query_param(dept)] if can_select_department(request.user) else []
    if dept and not can_select_department(request.user):
        params = [f'department={dept.pk}']
    if request.GET.get('batch'):
        params.append(f'batch={request.GET.get("batch")}')
    if request.GET.get('q'):
        params.append(f'q={request.GET.get("q")}')
    return f'{url}?{"&".join(params)}' if params else url


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def faculty_create(request):
    return faculty_list(request)


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def faculty_edit(request, pk):
    fac = get_object_or_404(Faculty, pk=pk)
    if request.method == 'POST':
        form = FacultyForm(request.POST, instance=fac)
        if form.is_valid():
            form.save()
            messages.success(request, 'Faculty updated successfully.')
            return redirect(_faculty_redirect(request, fac.department))
    form = FacultyForm(instance=fac)
    return render(request, 'portal/faculty/edit.html', {'form': form, 'faculty': fac})


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def faculty_delete(request, pk):
    fac = get_object_or_404(Faculty, pk=pk)
    dept = fac.department
    if request.method == 'POST':
        if fac.user:
            fac.user.delete()
        fac.delete()
        messages.success(request, 'Faculty deleted.')
    return redirect(_faculty_redirect(request, dept))


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def generate_faculty_credentials(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    faculty_ids = request.POST.getlist('faculty_ids')
    if not faculty_ids:
        messages.warning(request, 'Select at least one faculty member.')
        return redirect(_faculty_redirect(request, dept))

    creds = []
    for fid in faculty_ids:
        fac = get_object_or_404(Faculty, pk=fid)
        if dept and fac.department_id != dept.pk:
            continue
        if not dept and not _department_in_scope(request.user, ctx, fac.department):
            continue
        username = fac.mentor_code or f"fac_{fac.pk}"
        password = Faculty.generate_password()
        if fac.user:
            fac.user.set_password(password)
            fac.user.save()
            user = fac.user
        else:
            user = User.objects.create_user(
                username=username, password=password,
                role=User.Role.FACULTY, first_name=fac.name,
            )
            fac.user = user
            fac.credentials_generated = True
            fac.save()
        GeneratedCredential.objects.create(user=user, plain_password=password, generated_by=request.user)
        creds.append([user.username, password, fac.name, 'Faculty'])

    if request.POST.get('download') == '1':
        return export_credentials_excel(creds, 'faculty_credentials.xlsx')
    messages.success(request, f'Generated credentials for {len(creds)} faculty members.')
    return redirect(_faculty_redirect(request, dept))


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def student_list(request):
    ctx = get_user_context(request.user)
    is_super = request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN
    dept = resolve_department(request.user, ctx, request)
    departments = get_departments_for_user(request.user, ctx, request)
    batch_filter = request.GET.get('batch', '')
    search = request.GET.get('q', '').strip()

    form = StudentForm()
    upload_form = ExcelUploadForm()
    open_upload_modal = False

    if request.method == 'POST' and request.POST.get('form_type') == 'upload_students':
        upload_form = ExcelUploadForm(request.POST, request.FILES)
        target_dept = resolve_department(request.user, ctx, request)
        if request.user.is_superuser or request.user.role == User.Role.SEMESTER_ADMIN:
            dept_param = request.POST.get('department_id') or request.GET.get('department')
            if dept_param and dept_param not in ('all', ''):
                target_dept = get_object_or_404(Department, pk=dept_param)
        if not target_dept:
            messages.error(request, 'Please select a department first.')
            open_upload_modal = True
        elif upload_form.is_valid():
            created, updated, faculty_count = import_students_from_excel(
                upload_form.cleaned_data['file'], target_dept,
            )
            messages.success(
                request,
                f'Imported {created} new students, updated {updated}. '
                f'Found {faculty_count} mentor codes.',
            )
            return redirect(_student_redirect(request, target_dept))
        else:
            messages.error(request, 'Please upload a valid Excel or CSV file.')
            open_upload_modal = True
    elif request.method == 'POST' and request.POST.get('form_type') == 'add_student':
        form = StudentForm(request.POST)
        target_dept = resolve_department(request.user, ctx, request)
        if not target_dept:
            messages.error(request, 'Please select a department first.')
        elif form.is_valid():
            try:
                stu = form.save(commit=False)
                stu.department = target_dept
                stu.save()
                messages.success(request, f'Student "{stu.name}" added successfully.')
                return redirect(_student_redirect(request, target_dept))
            except IntegrityError:
                messages.error(request, 'A student with this enrollment number already exists.')
        else:
            messages.error(request, 'Please correct the errors below.')

    students = students_in_scope(request.user, ctx, dept).select_related('department')
    if batch_filter:
        students = students.filter(batch=batch_filter)
    if search:
        students = students.filter(
            Q(name__icontains=search) | Q(enrollment_no__icontains=search) |
            Q(roll_no__icontains=search) | Q(batch__icontains=search) |
            Q(mentor__icontains=search)
        )

    batches = students_in_scope(request.user, ctx, dept).values_list(
        'batch', flat=True
    ).distinct().order_by('batch')

    return render(request, 'portal/students/list.html', {
        'students': students,
        'form': form,
        'upload_form': upload_form,
        'open_upload_modal': open_upload_modal,
        'batches': batches,
        'batch_filter': batch_filter,
        'search': search,
        'total_count': students.count(),
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def student_create(request):
    return student_list(request)


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def student_edit(request, pk):
    stu = get_object_or_404(Student, pk=pk)
    if request.method == 'POST':
        form = StudentForm(request.POST, instance=stu)
        if form.is_valid():
            form.save()
            messages.success(request, 'Student updated successfully.')
            return redirect(_student_redirect(request, stu.department))
    form = StudentForm(instance=stu)
    return render(request, 'portal/students/edit.html', {'form': form, 'student': stu})


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def student_delete(request, pk):
    stu = get_object_or_404(Student, pk=pk)
    dept = stu.department
    if request.method == 'POST':
        if stu.user:
            stu.user.delete()
        stu.delete()
        messages.success(request, 'Student deleted.')
    return redirect(_student_redirect(request, dept))


def _resolve_student_password(post):
    """Return (password_or_none, error_message). None password means unique random per student."""
    mode = post.get('password_mode', 'random')
    if mode == 'default':
        default_password = post.get('default_password', '').strip()
        if not default_password:
            return None, 'Please enter a default password.'
        return default_password, None
    return None, None


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def generate_student_credentials(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    student_ids = request.POST.getlist('student_ids')
    batch = request.POST.get('batch', '')

    shared_password, password_error = _resolve_student_password(request.POST)
    if password_error:
        messages.error(request, password_error)
        return redirect(_student_redirect(request, dept))

    students = students_in_scope(request.user, ctx, dept)
    if student_ids:
        students = students.filter(pk__in=student_ids)
    elif batch:
        students = students.filter(batch=batch)

    creds = []
    for stu in students:
        username = stu.enrollment_no
        password = shared_password or Student.generate_password()
        if stu.user:
            stu.user.set_password(password)
            stu.user.save()
            user = stu.user
        else:
            user = User.objects.create_user(
                username=username, password=password,
                role=User.Role.STUDENT, first_name=stu.name,
            )
            stu.user = user
            stu.credentials_generated = True
            stu.save()
        GeneratedCredential.objects.create(user=user, plain_password=password, generated_by=request.user)
        creds.append([user.username, password, stu.name, 'Student'])

    if not creds:
        messages.warning(request, 'No students selected.')
        return redirect(_student_redirect(request, dept))

    if request.POST.get('download') == '1':
        return export_credentials_excel(creds, 'student_credentials.xlsx')
    if shared_password:
        messages.success(
            request,
            f'Set default password for {len(creds)} students.',
        )
    else:
        messages.success(request, f'Generated random credentials for {len(creds)} students.')
    return redirect(_student_redirect(request, dept))


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def export_students(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    batch = request.GET.get('batch', '')
    students = students_in_scope(request.user, ctx, dept)
    if batch:
        students = students.filter(batch=batch)
    return export_students_excel(students)


# ─── Attendance Sheets ───

@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def attendance_sheets(request):
    ensure_default_templates()
    ctx = get_user_context(request.user)
    is_super = request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN
    dept = resolve_department(request.user, ctx, request)
    departments = get_departments_for_user(request.user, ctx, request)
    semester = resolve_semester(request.user, ctx, request)

    if dept:
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
        templates = AttendanceSheetTemplate.objects.filter(
            Q(department=dept) | Q(department__isnull=True, semester=dept.semester) | Q(semester__isnull=True, department__isnull=True)
        )
    else:
        subjects = Subject.objects.filter(semester=semester) if semester else Subject.objects.all()
        templates = AttendanceSheetTemplate.objects.all()

    upload_form = AttendanceTemplateUploadForm()
    download_form = AttendanceDownloadForm()
    download_form.fields['subject'].queryset = subjects

    if dept:
        download_form.fields['semester_label'].initial = dept.sheet_semester_label or (dept.semester.name if dept.semester else '')
        download_form.fields['department_label'].initial = dept.sheet_department_label or dept.name
        batch_choices = [('', 'All Batches (GP)')]
        batch_choices += [
            (b, b) for b in Student.objects.filter(department=dept)
            .values_list('batch', flat=True).distinct().order_by('batch')
        ]
        download_form.fields['batch'].choices = batch_choices

    if request.method == 'POST':
        if request.POST.get('form_type') == 'upload_template':
            upload_form = AttendanceTemplateUploadForm(request.POST, request.FILES)
            if upload_form.is_valid():
                tmpl = upload_form.save(commit=False)
                tmpl.created_by = request.user
                if dept:
                    tmpl.department = dept
                    tmpl.semester = dept.semester
                elif semester:
                    tmpl.semester = semester
                tmpl.save()
                messages.success(request, f'Template "{tmpl.name}" uploaded for {tmpl.exam_type}.')
                return redirect(f'{reverse("portal:attendance_sheets")}?department={dept.pk}' if dept else reverse('portal:attendance_sheets'))
            messages.error(request, 'Please fix template upload errors.')

        elif request.POST.get('form_type') == 'download_sheet':
            download_form = AttendanceDownloadForm(request.POST)
            download_form.fields['subject'].queryset = subjects
            if dept:
                batch_choices = [('', 'All Batches (GP)')]
                batch_choices += [
                    (b, b) for b in Student.objects.filter(department=dept)
                    .values_list('batch', flat=True).distinct().order_by('batch')
                ]
                download_form.fields['batch'].choices = batch_choices
            if download_form.is_valid() and dept:
                exam_type = download_form.cleaned_data['exam_type']
                subject = download_form.cleaned_data['subject']
                semester_label = download_form.cleaned_data['semester_label']
                department_label = download_form.cleaned_data['department_label']
                batch_filter = download_form.cleaned_data.get('batch') or None

                dept.sheet_semester_label = semester_label
                dept.sheet_department_label = department_label
                dept.save(update_fields=['sheet_semester_label', 'sheet_department_label'])

                if exam_type == 'GP':
                    return generate_attendance_workbook(
                        None, dept, subject, exam_type,
                        semester_label, department_label,
                        batch_filter=batch_filter,
                    )

                tmpl = templates.filter(exam_type=exam_type).order_by('-created_at').first()
                template_path = tmpl.template_file.path if tmpl and tmpl.template_file else get_default_template_path(exam_type)
                if not template_path or not os.path.exists(template_path):
                    messages.error(request, f'No {exam_type} template found. Please upload one first.')
                else:
                    return generate_attendance_workbook(
                        template_path, dept, subject, exam_type,
                        semester_label, department_label,
                        batch_filter=batch_filter,
                    )
            else:
                messages.error(request, 'Select department, subject and fill all fields.')

    return render(request, 'portal/attendance/sheets.html', {
        'upload_form': upload_form,
        'download_form': download_form,
        'templates': templates,
        'subjects': subjects,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def attendance_template_delete(request, pk):
    tmpl = get_object_or_404(AttendanceSheetTemplate, pk=pk)
    dept = tmpl.department
    if request.method == 'POST':
        if tmpl.template_file:
            tmpl.template_file.delete(save=False)
        tmpl.delete()
        messages.success(request, 'Attendance template deleted.')
    url = reverse('portal:attendance_sheets')
    if dept:
        return redirect(f'{url}?department={dept.pk}')
    return redirect(url)


# ─── Marksheet Templates ───

def _marksheet_templates_queryset(user, ctx, dept=None):
    if dept:
        return MarksheetTemplate.objects.filter(
            Q(department=dept) | Q(department__isnull=True, semester=dept.semester)
        ).select_related('subject', 'department', 'semester')
    semester = ctx.get('semester') or resolve_semester(user, ctx)
    if semester:
        return MarksheetTemplate.objects.filter(
            Q(semester=semester) | Q(department__semester=semester)
        ).select_related('subject', 'department', 'semester')
    return MarksheetTemplate.objects.all().select_related('subject', 'department', 'semester')


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def marksheet_templates(request):
    ensure_default_marksheet_template()
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    semester = resolve_semester(request.user, ctx, request)

    if dept:
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
    elif semester:
        subjects = Subject.objects.filter(semester=semester)
    else:
        subjects = Subject.objects.all()

    templates = _marksheet_templates_queryset(request.user, ctx, dept)
    upload_form = MarksheetTemplateUploadForm()
    upload_form.fields['subject'].queryset = subjects
    download_form = MarksheetDownloadForm()
    download_form.fields['subject'].queryset = subjects

    if dept:
        download_form.fields['semester_label'].initial = dept.sheet_semester_label or (dept.semester.name if dept.semester else '')
        download_form.fields['department_label'].initial = dept.sheet_department_label or dept.name

    filter_exam = request.GET.get('exam_type', '')
    if filter_exam:
        templates = templates.filter(exam_type=filter_exam)

    if request.method == 'POST':
        if request.POST.get('form_type') == 'upload_template':
            upload_form = MarksheetTemplateUploadForm(request.POST, request.FILES)
            upload_form.fields['subject'].queryset = subjects
            if upload_form.is_valid():
                exam_type = upload_form.cleaned_data['exam_type']
                subject = upload_form.cleaned_data['subject']
                replace = upload_form.cleaned_data.get('replace_existing')
                tmpl_semester = dept.semester if dept else (subject.semester or semester)
                tmpl_department = dept if request.user.role == User.Role.DEPARTMENT_ADMIN else None
                if not tmpl_semester:
                    messages.error(request, 'Could not determine semester for this template.')
                else:
                    existing = MarksheetTemplate.objects.filter(
                        exam_type=exam_type,
                        subject=subject,
                        semester=tmpl_semester,
                        department=tmpl_department,
                    ).first()
                    if existing and replace:
                        if existing.template_file:
                            existing.template_file.delete(save=False)
                        existing.delete()

                    tmpl = upload_form.save(commit=False)
                    tmpl.semester = tmpl_semester
                    tmpl.department = tmpl_department
                    tmpl.created_by = request.user
                    tmpl.save()
                    tmpl.schema = parse_marksheet_template(tmpl.template_file.path)
                    tmpl.save(update_fields=['schema'])
                    scope = 'all departments' if not tmpl_department else tmpl_department.name
                    messages.success(request, f'Marksheet template saved for {subject.name} ({exam_type}) — {scope}.')
                    url = reverse('portal:marksheet_templates')
                    if dept:
                        return redirect(f'{url}?department={dept.pk}&exam_type={exam_type}')
                    return redirect(f'{url}?exam_type={exam_type}')
            messages.error(request, 'Please fix template upload errors.')

        elif request.POST.get('form_type') == 'download_marksheet':
            download_form = MarksheetDownloadForm(request.POST)
            download_form.fields['subject'].queryset = subjects
            if download_form.is_valid() and dept:
                exam_type = download_form.cleaned_data['exam_type']
                subject = download_form.cleaned_data['subject']
                semester_label = download_form.cleaned_data['semester_label']
                department_label = download_form.cleaned_data['department_label']

                dept.sheet_semester_label = semester_label
                dept.sheet_department_label = department_label
                dept.save(update_fields=['sheet_semester_label', 'sheet_department_label'])

                tmpl = resolve_marksheet_template(exam_type, subject, dept)
                if not tmpl:
                    messages.error(request, f'No {exam_type} marksheet template uploaded for {subject.name}. Please upload first.')
                else:
                    return generate_marksheet_workbook(tmpl, dept, semester_label, department_label)
            else:
                messages.error(request, 'Select department, exam type, subject and fill all fields.')

    return render(request, 'portal/marksheet/templates.html', {
        'upload_form': upload_form,
        'download_form': download_form,
        'templates': templates,
        'subjects': subjects,
        'filter_exam': filter_exam,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def marksheet_template_delete(request, pk):
    tmpl = get_object_or_404(MarksheetTemplate, pk=pk)
    dept = tmpl.department
    exam_type = tmpl.exam_type
    if request.method == 'POST':
        if tmpl.template_file:
            tmpl.template_file.delete(save=False)
        tmpl.delete()
        messages.success(request, 'Marksheet template deleted.')
    url = reverse('portal:marksheet_templates')
    if dept:
        return redirect(f'{url}?department={dept.pk}&exam_type={exam_type}')
    return redirect(f'{url}?exam_type={exam_type}')


def _group_faculty_duties(duties_qs):
    """Merge internal/external rows for the same duty slot into one display row."""
    groups = {}
    order = []
    for duty in duties_qs:
        key = (
            duty.duty_date,
            duty.subject_id,
            duty.exam_type,
            duty.batch,
            duty.time_slot or '',
            duty.room_no or '',
        )
        if key not in groups:
            groups[key] = {
                'duty_date': duty.duty_date,
                'subject': duty.subject,
                'exam_type': duty.exam_type,
                'batch': duty.batch,
                'time_slot': duty.time_slot,
                'room_no': duty.room_no,
                'internal': None,
                'external': None,
            }
            order.append(key)
        if duty.duty_role == FacultyDutyAssignment.DutyRole.INTERNAL:
            groups[key]['internal'] = duty
        else:
            groups[key]['external'] = duty
    return [groups[key] for key in order]


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def faculty_duty_list(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    duties = FacultyDutyAssignment.objects.select_related(
        'faculty', 'subject', 'department',
    ).filter(is_active=True).order_by('-duty_date', 'batch', 'duty_role')
    if dept:
        duties = duties.filter(department=dept)
    elif ctx.get('semester'):
        duties = duties.filter(department__semester=ctx['semester'])

    subjects = Subject.objects.all()
    faculty_qs = Faculty.objects.all()
    if dept:
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
        faculty_qs = Faculty.objects.filter(department=dept)
        batches = department_batches(dept)
    else:
        batches = Student.objects.values_list('batch', flat=True).distinct().order_by('batch')

    duty_form = FacultyDutyAssignmentForm()
    upload_form = DutyScheduleUploadForm()
    duty_form.fields['subject'].queryset = subjects
    duty_form.fields['faculty'].queryset = faculty_qs
    duty_form.fields['batch'].choices = [(b, b) for b in batches]

    external_without_login = Faculty.objects.none()
    recent_uploads = DutyScheduleUpload.objects.none()
    last_import_summary = None
    if dept:
        external_without_login = Faculty.objects.filter(
            department=dept, is_external=True, user__isnull=True,
        )
        recent_uploads = DutyScheduleUpload.objects.filter(department=dept)[:5]
        last_upload = recent_uploads.first()
        if last_upload and last_upload.summary:
            last_import_summary = last_upload.summary

    if request.method == 'POST':
        form_type = request.POST.get('form_type')

        if form_type == 'upload_schedule':
            if not dept:
                messages.error(request, 'Select a department first.')
            else:
                upload_form = DutyScheduleUploadForm(request.POST, request.FILES)
                if upload_form.is_valid():
                    try:
                        _, summary = import_duty_schedule(
                            dept,
                            upload_form.cleaned_data['schedule_file'],
                            request.user,
                            replace_existing=upload_form.cleaned_data.get('replace_existing'),
                        )
                        msg = (
                            f"Schedule imported: {summary['created']} created, "
                            f"{summary['updated']} updated"
                        )
                        if summary.get('external_created'):
                            msg += f", {summary['external_created']} new external examiner(s)"
                        messages.success(request, msg)
                        if summary.get('warnings'):
                            messages.warning(
                                request,
                                f"{len(summary['warnings'])} warning(s) — see import summary below.",
                            )
                    except Exception as exc:
                        messages.error(request, f'Import failed: {exc}')
                    return redirect(f'{reverse("portal:faculty_duty_list")}?department={dept.pk}')
                messages.error(request, 'Please upload a valid Excel schedule file.')

        elif form_type == 'assign_duty':
            if not dept:
                messages.error(request, 'Select a department first.')
            else:
                duty_form = FacultyDutyAssignmentForm(request.POST)
                duty_form.fields['subject'].queryset = subjects
                duty_form.fields['faculty'].queryset = faculty_qs
                duty_form.fields['batch'].choices = [(b, b) for b in batches]
                if duty_form.is_valid():
                    duty = duty_form.save(commit=False)
                    duty.department = dept
                    duty.assigned_by = request.user
                    try:
                        duty.save()
                        messages.success(request, f'Duty assigned to {duty.faculty.name} for {duty.batch}.')
                    except IntegrityError:
                        messages.error(request, 'This faculty already has this duty assignment for that date.')
                    return redirect(f'{reverse("portal:faculty_duty_list")}?department={dept.pk}')
                messages.error(request, 'Please fix duty assignment errors.')

        elif form_type == 'generate_external_credentials' and dept:
            creds = []
            for fac in external_without_login:
                username = fac.mentor_code or f'ext_{fac.pk}'
                password = Faculty.generate_password()
                user = User.objects.create_user(
                    username=username, password=password,
                    role=User.Role.FACULTY, first_name=fac.name,
                )
                fac.user = user
                fac.credentials_generated = True
                fac.save()
                GeneratedCredential.objects.create(
                    user=user, plain_password=password, generated_by=request.user,
                )
                creds.append([user.username, password, fac.name, 'External Examiner'])
            if request.POST.get('download') == '1' and creds:
                return export_credentials_excel(creds, 'external_examiner_credentials.xlsx')
            if creds:
                messages.success(request, f'Generated login for {len(creds)} external examiner(s).')
            else:
                messages.info(request, 'All external examiners already have login credentials.')
            return redirect(f'{reverse("portal:faculty_duty_list")}?department={dept.pk}')

    return render(request, 'portal/duty/list.html', {
        'duty_rows': _group_faculty_duties(list(duties)),
        'duty_form': duty_form,
        'upload_form': upload_form,
        'external_without_login': external_without_login,
        'recent_uploads': recent_uploads,
        'last_import_summary': last_import_summary,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def faculty_duty_delete(request, pk):
    duty = get_object_or_404(FacultyDutyAssignment, pk=pk)
    ctx = get_user_context(request.user)
    if not _department_in_scope(request.user, ctx, duty.department):
        messages.error(request, 'You do not have permission to remove this duty.')
        return redirect('portal:faculty_duty_list')
    dept = duty.department
    if request.method == 'POST':
        duty.is_active = False
        duty.save(update_fields=['is_active'])
        messages.success(request, 'Duty assignment removed.')
    url = reverse('portal:faculty_duty_list')
    if dept:
        return redirect(f'{url}?department={dept.pk}')
    return redirect(url)


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def final_marksheet_download(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    download_form = FinalMarksheetDownloadForm()
    if dept:
        download_form.fields['subject'].queryset = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
        download_form.fields['semester_label'].initial = dept.sheet_semester_label or (dept.semester.name if dept.semester else '')
        download_form.fields['department_label'].initial = dept.sheet_department_label or dept.name

    if request.method == 'POST' and dept:
        download_form = FinalMarksheetDownloadForm(request.POST)
        download_form.fields['subject'].queryset = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
        if download_form.is_valid():
            exam_type = download_form.cleaned_data['exam_type']
            subject = download_form.cleaned_data['subject']
            semester_label = download_form.cleaned_data['semester_label']
            department_label = download_form.cleaned_data['department_label']
            action = request.POST.get('action', 'filled')
            dept.sheet_semester_label = semester_label
            dept.sheet_department_label = department_label
            dept.save(update_fields=['sheet_semester_label', 'sheet_department_label'])
            tmpl = resolve_marksheet_template(exam_type, subject, dept)
            if action == 'compiled':
                return generate_compiled_marksheet_workbook(
                    exam_type, subject, dept, semester_label, department_label, template=tmpl,
                )
            combined = download_form.cleaned_data.get('download_type') == 'combined'
            if not tmpl:
                messages.error(request, f'No {exam_type} marksheet template for {subject.name}. Upload template first.')
            else:
                return generate_marksheet_workbook(
                    tmpl, dept, semester_label, department_label, combined=combined,
                )
        else:
            messages.error(request, 'Please check the form — some fields are missing or invalid.')

    return render(request, 'portal/marksheet/final_download.html', {
        'download_form': download_form,
        **dept_filter_context(request.user, ctx, request, dept),
    })


def _subjects_for_department(dept):
    if not dept:
        return Subject.objects.none()
    return Subject.objects.filter(
        Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
    )


def _syllabus_upload_department(user, dept):
    if user.role == User.Role.DEPARTMENT_ADMIN:
        return dept
    if user.role == User.Role.SEMESTER_ADMIN:
        return None
    return dept


def _syllabus_papers_queryset(model, user, ctx, dept):
    qs = model.objects.filter(is_active=True).select_related('subject', 'department', 'semester')
    if dept:
        return qs.filter(semester=dept.semester).filter(
            Q(department=dept) | Q(department__isnull=True)
        )
    semester = ctx.get('semester') or resolve_semester(user, ctx)
    if semester:
        return qs.filter(semester=semester)
    return qs


def _faculty_syllabus_papers_queryset(model, faculty_dept):
    return model.objects.filter(
        is_active=True,
        semester=faculty_dept.semester,
    ).filter(
        Q(department=faculty_dept) | Q(department__isnull=True)
    ).select_related('subject', 'department')


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def syllabus_papers_manage(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    if dept:
        subjects = _subjects_for_department(dept)
    else:
        semester = ctx.get('semester') or resolve_semester(request.user, ctx, request)
        subjects = Subject.objects.filter(semester=semester) if semester else Subject.objects.all()

    syllabi = _syllabus_papers_queryset(SubjectSyllabus, request.user, ctx, dept)
    papers = _syllabus_papers_queryset(SubjectPaper, request.user, ctx, dept)

    syllabus_form = SyllabusUploadForm()
    paper_form = PaperUploadForm()
    syllabus_form.fields['subject'].queryset = subjects
    paper_form.fields['subject'].queryset = subjects

    if request.method == 'POST':
        form_type = request.POST.get('form_type')
        scope_dept = _syllabus_upload_department(request.user, dept)
        upload_semester = dept.semester if dept else (ctx.get('semester') or resolve_semester(request.user, ctx, request))

        if form_type == 'upload_syllabus':
            if not upload_semester:
                messages.error(request, 'Select a department or semester first.')
            else:
                syllabus_form = SyllabusUploadForm(request.POST, request.FILES)
                syllabus_form.fields['subject'].queryset = subjects
                if syllabus_form.is_valid():
                    subject = syllabus_form.cleaned_data['subject']
                    replace = syllabus_form.cleaned_data.get('replace_existing')
                    existing = SubjectSyllabus.objects.filter(
                        subject=subject, semester=upload_semester, department=scope_dept, is_active=True,
                    ).first()
                    if existing and replace:
                        if existing.file:
                            existing.file.delete(save=False)
                        existing.delete()
                    elif existing:
                        messages.error(request, 'Syllabus already exists for this subject. Check replace option.')
                        return redirect(_syllabus_redirect_url(dept))

                    obj = syllabus_form.save(commit=False)
                    obj.semester = upload_semester
                    obj.department = scope_dept
                    obj.created_by = request.user
                    obj.save()
                    messages.success(request, f'Syllabus uploaded for {subject.name}.')
                    return redirect(_syllabus_redirect_url(dept))
                messages.error(request, 'Please fix syllabus upload errors.')

        elif form_type == 'upload_paper':
            if not upload_semester:
                messages.error(request, 'Select a department or semester first.')
            else:
                paper_form = PaperUploadForm(request.POST, request.FILES)
                paper_form.fields['subject'].queryset = subjects
                if paper_form.is_valid():
                    subject = paper_form.cleaned_data['subject']
                    phase = paper_form.cleaned_data['phase']
                    replace = paper_form.cleaned_data.get('replace_existing')
                    existing = SubjectPaper.objects.filter(
                        subject=subject, semester=upload_semester,
                        department=scope_dept, phase=phase, is_active=True,
                    ).first()
                    if existing and replace:
                        if existing.file:
                            existing.file.delete(save=False)
                        existing.delete()
                    elif existing:
                        messages.error(request, f'Phase {phase} paper already exists. Check replace option.')
                        return redirect(_syllabus_redirect_url(dept))

                    obj = paper_form.save(commit=False)
                    obj.semester = upload_semester
                    obj.department = scope_dept
                    obj.created_by = request.user
                    try:
                        obj.save()
                        messages.success(request, f'Phase {phase} paper uploaded for {subject.name}.')
                    except IntegrityError:
                        messages.error(request, f'Phase {phase} paper already exists for {subject.name}.')
                    return redirect(_syllabus_redirect_url(dept))
                messages.error(request, 'Please fix paper upload errors.')

    return render(request, 'portal/syllabus/manage.html', {
        'syllabi': syllabi,
        'papers': papers,
        'syllabus_form': syllabus_form,
        'paper_form': paper_form,
        **dept_filter_context(request.user, ctx, request, dept),
    })


def _syllabus_redirect_url(dept):
    url = reverse('portal:syllabus_papers_manage')
    if dept:
        return f'{url}?department={dept.pk}'
    return url


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def syllabus_delete(request, pk):
    obj = get_object_or_404(SubjectSyllabus, pk=pk)
    dept = obj.department
    if request.method == 'POST':
        obj.is_active = False
        obj.save(update_fields=['is_active'])
        messages.success(request, 'Syllabus removed.')
    return redirect(_syllabus_redirect_url(dept))


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def paper_delete(request, pk):
    obj = get_object_or_404(SubjectPaper, pk=pk)
    dept = obj.department
    if request.method == 'POST':
        obj.is_active = False
        obj.save(update_fields=['is_active'])
        messages.success(request, 'Paper removed.')
    return redirect(_syllabus_redirect_url(dept))


@role_required(User.Role.FACULTY)
def faculty_syllabus(request):
    faculty = request.user.faculty_profile
    dept = faculty.department
    syllabi = _faculty_syllabus_papers_queryset(SubjectSyllabus, dept).order_by('subject__name')
    subject_filter = request.GET.get('subject', '')
    if subject_filter:
        syllabi = syllabi.filter(subject_id=subject_filter)
    subjects = _subjects_for_department(dept).filter(
        pk__in=_faculty_syllabus_papers_queryset(SubjectSyllabus, dept).values_list('subject_id', flat=True)
    )
    return render(request, 'portal/faculty/syllabus.html', {
        'syllabi': syllabi,
        'subjects': subjects,
        'subject_filter': subject_filter,
        'department': dept,
    })


@role_required(User.Role.FACULTY)
def faculty_papers(request):
    faculty = request.user.faculty_profile
    dept = faculty.department
    papers = _faculty_syllabus_papers_queryset(SubjectPaper, dept).order_by('subject__name', 'phase')
    subject_filter = request.GET.get('subject', '')
    if subject_filter:
        papers = papers.filter(subject_id=subject_filter)
    subjects = _subjects_for_department(dept).filter(
        pk__in=_faculty_syllabus_papers_queryset(SubjectPaper, dept).values_list('subject_id', flat=True)
    )
    return render(request, 'portal/faculty/papers.html', {
        'papers': papers,
        'subjects': subjects,
        'subject_filter': subject_filter,
        'department': dept,
    })


def _faculty_can_access_document(faculty_dept, obj):
    if not obj.is_active:
        return False
    if obj.semester_id != faculty_dept.semester_id:
        return False
    if obj.department_id and obj.department_id != faculty_dept.pk:
        return False
    return True


def _document_content_type(filename):
    low = (filename or '').lower()
    if low.endswith('.pdf'):
        return 'application/pdf'
    if low.endswith('.docx'):
        return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    if low.endswith('.doc'):
        return 'application/msword'
    guessed, _ = mimetypes.guess_type(filename)
    return guessed or 'application/octet-stream'


def _serve_faculty_document(request, obj, faculty_dept):
    if not _faculty_can_access_document(faculty_dept, obj):
        return HttpResponseForbidden('Access denied.')
    filename = obj.file.name.split('/')[-1]
    inline = request.GET.get('download') != '1'
    response = FileResponse(
        obj.file.open('rb'),
        content_type=_document_content_type(filename),
        as_attachment=not inline,
        filename=filename,
    )
    if inline:
        response['Content-Disposition'] = f'inline; filename="{filename}"'
    return response


def _faculty_document_view_context(request, obj, doc_type, back_url_name):
    dept = request.user.faculty_profile.department
    if not _faculty_can_access_document(dept, obj):
        messages.error(request, 'You do not have access to this document.')
        return None
    filename = obj.file.name.lower()
    is_pdf = filename.endswith('.pdf')
    file_pk = obj.pk
    if doc_type == 'paper':
        file_url = reverse('portal:faculty_paper_file', kwargs={'pk': file_pk})
    else:
        file_url = reverse('portal:faculty_syllabus_file', kwargs={'pk': file_pk})
    return {
        'document': obj,
        'doc_type': doc_type,
        'title': obj.display_title,
        'file_url': file_url,
        'download_url': f'{file_url}?download=1',
        'back_url': reverse(back_url_name),
        'is_pdf': is_pdf,
        'is_word': filename.endswith('.doc') or filename.endswith('.docx'),
        'is_docx': filename.endswith('.docx'),
    }


@role_required(User.Role.FACULTY)
def faculty_paper_view(request, pk):
    paper = get_object_or_404(SubjectPaper, pk=pk, is_active=True)
    ctx = _faculty_document_view_context(request, paper, 'paper', 'portal:faculty_papers')
    if ctx is None:
        return redirect('portal:faculty_papers')
    return render(request, 'portal/faculty/document_view.html', ctx)


@role_required(User.Role.FACULTY)
@xframe_options_sameorigin
def faculty_paper_file(request, pk):
    paper = get_object_or_404(SubjectPaper, pk=pk, is_active=True)
    return _serve_faculty_document(request, paper, request.user.faculty_profile.department)


@role_required(User.Role.FACULTY)
def faculty_syllabus_view(request, pk):
    syllabus = get_object_or_404(SubjectSyllabus, pk=pk, is_active=True)
    ctx = _faculty_document_view_context(request, syllabus, 'syllabus', 'portal:faculty_syllabus')
    if ctx is None:
        return redirect('portal:faculty_syllabus')
    return render(request, 'portal/faculty/document_view.html', ctx)


@role_required(User.Role.FACULTY)
@xframe_options_sameorigin
def faculty_syllabus_file(request, pk):
    syllabus = get_object_or_404(SubjectSyllabus, pk=pk, is_active=True)
    return _serve_faculty_document(request, syllabus, request.user.faculty_profile.department)


# ─── Form Templates ───

def _templates_for_user(user, ctx):
    if user.role == User.Role.DEPARTMENT_ADMIN:
        dept = ctx['department']
        return FormTemplate.objects.filter(
            Q(department=dept) | Q(department__isnull=True, semester=dept.semester)
        )
    semester = ctx.get('semester') or resolve_semester(user, ctx)
    if semester:
        return FormTemplate.objects.filter(semester=semester)
    return FormTemplate.objects.all()


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def form_template_list(request):
    ctx = get_user_context(request.user)
    templates = _templates_for_user(request.user, ctx).prefetch_related('fields')
    return render(request, 'portal/forms/template_list.html', {'templates': templates})


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def form_template_create(request):
    ctx = get_user_context(request.user)
    if request.method == 'POST':
        form = FormTemplateForm(request.POST, request.FILES)
        if form.is_valid():
            tmpl = form.save(commit=False)
            if request.user.role == User.Role.DEPARTMENT_ADMIN:
                tmpl.department = ctx['department']
                tmpl.semester = ctx['department'].semester
            else:
                tmpl.semester = ctx.get('semester') or resolve_semester(request.user, ctx, request)
            if not tmpl.semester:
                messages.error(request, 'No semester context. Please select a semester first.')
                return render(request, 'portal/forms/template_create.html', {'form': form})
            tmpl.created_by = request.user
            tmpl.save()
            if 'file' in request.FILES:
                tmpl.uploaded_file = request.FILES['file']
                tmpl.save()
                count = create_form_fields_from_excel(
                    request.FILES['file'], tmpl,
                    semester=tmpl.semester, department=tmpl.department,
                )
                messages.success(request, f'Template created — {count} field(s) imported from Excel.')
            else:
                messages.success(request, 'Template created. Upload Excel or add fields manually.')
            return redirect('portal:form_template_fields', pk=tmpl.pk)
    else:
        form = FormTemplateForm()
    return render(request, 'portal/forms/template_create.html', {'form': form})


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def form_template_delete(request, pk):
    tmpl = get_object_or_404(FormTemplate, pk=pk)
    if request.method == 'POST':
        tmpl.delete()
        messages.success(request, 'Template deleted.')
    return redirect('portal:form_template_list')


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def form_template_fields(request, pk):
    ctx = get_user_context(request.user)
    tmpl = get_object_or_404(FormTemplate, pk=pk)

    if request.method == 'POST' and request.POST.get('form_type') == 'import_excel':
        upload = request.FILES.get('file')
        if not upload:
            messages.error(request, 'Please select an Excel file.')
        else:
            replace = request.POST.get('import_mode') != 'append'
            count, parsed = import_fields_from_excel(upload, tmpl, replace=replace)
            tmpl.uploaded_file = upload
            tmpl.save()
            case_count = import_cases_from_excel(upload, tmpl.semester, tmpl.department)
            if count:
                messages.success(request, f'Imported {count} field(s) from Excel.')
            else:
                messages.warning(request, 'No extra fields found in Excel (built-in fields like Case, Title, Enrollment are skipped).')
            if case_count:
                messages.success(request, f'Also imported {case_count} case(s) from Excel.')

    if request.method == 'POST' and request.POST.get('form_type') == 'add_field':
        form = FormFieldForm(request.POST)
        if form.is_valid():
            field = form.save(commit=False)
            field.template = tmpl
            from .gp_utils import slugify_field_name
            field.field_name = slugify_field_name(field.field_label)
            field.save()
            messages.success(request, f'Field "{field.field_label}" added.')
            return redirect('portal:form_template_fields', pk=pk)
    else:
        form = FormFieldForm()

    if request.method == 'POST' and request.POST.get('form_type') == 'delete_field':
        fid = request.POST.get('field_id')
        FormField.objects.filter(pk=fid, template=tmpl).delete()
        messages.success(request, 'Field removed.')

    group_fields = tmpl.fields.filter(field_scope=FormField.FieldScope.GROUP)
    member_fields = tmpl.fields.filter(field_scope=FormField.FieldScope.MEMBER)
    return render(request, 'portal/forms/template_fields.html', {
        'template': tmpl,
        'form': form,
        'group_fields': group_fields,
        'member_fields': member_fields,
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def project_case_list(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request) if request.user.role != User.Role.DEPARTMENT_ADMIN else ctx.get('department')
    semester = dept.semester if dept else ctx.get('semester') or resolve_semester(request.user, ctx, request)

    if request.method == 'POST' and request.POST.get('form_type') == 'add_case':
        form = ProjectCaseForm(request.POST)
        if form.is_valid() and semester:
            case = form.save(commit=False)
            case.semester = semester
            if request.user.role == User.Role.DEPARTMENT_ADMIN:
                case.department = ctx['department']
            elif dept:
                case.department = dept
            case.save()
            messages.success(request, f'Case "{case.name}" added.')
            return redirect('portal:project_case_list')
    else:
        form = ProjectCaseForm()

    if request.method == 'POST' and request.POST.get('form_type') == 'delete_case':
        ProjectCase.objects.filter(pk=request.POST.get('case_id')).delete()
        messages.success(request, 'Case deleted.')

    cases = ProjectCase.objects.filter(semester=semester) if semester else ProjectCase.objects.none()
    scope_dept = None
    if request.user.role == User.Role.DEPARTMENT_ADMIN:
        scope_dept = ctx.get('department')
    elif dept:
        scope_dept = dept
    if scope_dept:
        cases = cases.filter(Q(department=scope_dept) | Q(department__isnull=True))

    return render(request, 'portal/forms/project_cases.html', {
        'cases': cases,
        'form': form,
        'semester': semester,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def gp_deadline_settings(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    if request.user.role == User.Role.DEPARTMENT_ADMIN:
        dept = ctx.get('department')

    if not dept:
        messages.error(request, 'Please select a department to configure the GP submission deadline.')
        return redirect('portal:user_management')

    if request.method == 'POST':
        form = GPDeadlineForm(request.POST, department=dept)
        if form.is_valid():
            deadline = form.cleaned_data['gp_submission_deadline']
            if deadline and timezone.is_naive(deadline):
                deadline = timezone.make_aware(deadline, timezone.get_current_timezone())
            dept.gp_submission_deadline = deadline
            dept.save(update_fields=['gp_submission_deadline'])
            if deadline:
                local = timezone.localtime(deadline)
                messages.success(
                    request,
                    f'GP submission deadline set to {local.strftime("%d %b %Y, %I:%M %p")}.',
                )
            else:
                messages.success(request, 'GP submission deadline removed. Students can submit and edit anytime.')
            return redirect(f'{reverse("portal:gp_deadline_settings")}?department={dept.pk}')
    else:
        form = GPDeadlineForm(department=dept)

    return render(request, 'portal/forms/gp_deadline.html', {
        'form': form,
        'department': dept,
        'gp_submission_locked': is_gp_submission_locked(dept),
        **dept_filter_context(request.user, ctx, request, dept),
    })


# ─── Student: GP & Profile ───

@login_required
def student_profile(request):
    if request.user.role != User.Role.STUDENT:
        return redirect('portal:dashboard')
    student = request.user.student_profile
    return render(request, 'portal/student/profile.html', {'student': student})


def _gp_student_context(student, editing_group=None, post_data=None):
    dept = student.department
    batch_students = Student.objects.filter(department=dept, batch=student.batch).order_by('roll_no')
    subjects = subjects_for_department(dept)
    template = get_gp_template(dept)
    project_cases = get_project_cases(dept)
    existing_groups = GPGroup.objects.filter(members=student).select_related(
        'subject', 'leader', 'department', 'project_case'
    ).prefetch_related('member_details__student', 'members').order_by('-updated_at')

    bundle_groups = list(get_bundle_groups(editing_group)) if editing_group else []
    bundle_entries = bundle_subject_entries_from_groups(bundle_groups) if bundle_groups else []

    students_json = [
        {
            'pk': s.pk,
            'name': s.name,
            'enrollment_no': s.enrollment_no,
            'roll_no': s.roll_no,
            'branch': s.branch,
            'batch': s.batch,
        }
        for s in batch_students
    ]

    member_details = {}
    ref_group = bundle_groups[0] if bundle_groups else editing_group
    if ref_group:
        for d in ref_group.member_details.all():
            member_details[str(d.student_id)] = {
                'gender': d.gender,
                'member_data': d.member_data,
            }

    initial_members = (
        list(ref_group.member_details.values_list('student_id', flat=True))
        if ref_group else [student.pk]
    )
    exclude_group_ids = [g.pk for g in bundle_groups] if bundle_groups else (
        [editing_group.pk] if editing_group else []
    )
    is_group_leader = GPGroup.objects.filter(leader=student).exists()
    gp_locked = is_gp_submission_locked(dept)
    show_submission_form = (
        editing_group is not None
        or not existing_groups.exists()
        or is_group_leader
    ) and not gp_locked

    subject_options = build_gp_subject_selection_options(subjects)
    initial_selection = selection_value_for_groups(bundle_groups) if bundle_groups else ''

    if post_data:
        repost_sel = post_data.get('subject_selection', '').strip()
        if repost_sel:
            initial_selection = repost_sel
            repost_ids = parse_subject_selection(repost_sel)
            case_bundle = post_data.get('case_bundle', '').strip()
            title_combined = post_data.get('title_combined', '').strip()
            bundle_entries = []
            for sid in repost_ids:
                subj = next((s for s in subjects if s.pk == sid), None)
                case_raw, title = _resolve_subject_entry_from_post(
                    post_data, sid, subj,
                    case_bundle=case_bundle,
                    title_combined=title_combined,
                )
                bundle_entries.append({
                    'subject_id': sid,
                    'subject_name': subj.name if subj else '',
                    'case_id': int(case_raw) if str(case_raw).isdigit() else None,
                    'title': title,
                })
        repost_members = []
        for mid in post_data.getlist('member_ids'):
            if str(mid).isdigit():
                repost_members.append(int(mid))
        if repost_members:
            initial_members = repost_members

    taken_titles_map = get_taken_titles_by_subject(
        dept, student.batch, subjects, exclude_group_ids=exclude_group_ids,
    )
    classmate_titles_by_subject = [
        {'subject': s, 'titles': taken_titles_map.get(str(s.pk), [])}
        for s in subjects
    ]

    repost_values = {}
    if post_data:
        repost_values = {
            'gender_diversity': post_data.get('gender_diversity', ''),
            'religion_diversity': post_data.get('religion_diversity', ''),
        }

    return {
        'student': student,
        'batch_students': batch_students,
        'students_json': json.dumps(students_json),
        'subjects': subjects,
        'subject_options': subject_options,
        'subject_options_json': json.dumps(subject_options),
        'project_cases': project_cases,
        'template': template,
        'group_fields': get_group_fields(template),
        'member_fields': get_member_fields(template),
        'existing_groups': existing_groups,
        'editing_group': editing_group,
        'bundle_groups': bundle_groups,
        'bundle_entries_json': json.dumps(bundle_entries),
        'initial_subject_selection': initial_selection,
        'member_details_json': json.dumps(member_details),
        'group_data_json': json.dumps(ref_group.group_data if ref_group else {}),
        'initial_members_json': json.dumps(initial_members),
        'faculty_by_subject_json': json.dumps(
            get_faculty_map_for_subjects(dept, student.batch, subjects),
        ),
        'taken_by_subject_json': json.dumps(
            get_taken_members_by_subject(dept, subjects, exclude_group_ids=exclude_group_ids),
        ),
        'taken_titles_by_subject_json': json.dumps(taken_titles_map),
        'classmate_titles_by_subject': classmate_titles_by_subject,
        'repost_values': repost_values,
        'exclude_group_ids_json': json.dumps(exclude_group_ids),
        'show_submission_form': show_submission_form,
        'is_group_leader': is_group_leader,
        'gp_submission_locked': gp_locked,
        'gp_submission_deadline': dept.gp_submission_deadline,
        'gender_choices': GENDER_CHOICES,
        'yes_no_choices': YES_NO_CHOICES,
        'max_group_members': MAX_GP_GROUP_MEMBERS,
    }


@login_required
def gp_project(request):
    if request.user.role != User.Role.STUDENT:
        return redirect('portal:dashboard')
    student = request.user.student_profile
    dept = student.department

    if request.method == 'POST' and request.POST.get('form_type') == 'create_group':
        group, errors = save_gp_submission(student, dept, request.POST)
        if group:
            messages.success(request, 'GP project submitted successfully!')
            return redirect('portal:gp_project')
        for err in errors:
            messages.error(request, err)
        ctx = _gp_student_context(student, post_data=request.POST)
        return render(request, 'portal/student/gp_project.html', ctx)

    ctx = _gp_student_context(student)
    return render(request, 'portal/student/gp_project.html', ctx)


@login_required
def gp_group_edit(request, group_pk):
    if request.user.role != User.Role.STUDENT:
        return redirect('portal:dashboard')
    student = request.user.student_profile
    group = get_object_or_404(GPGroup, pk=group_pk)
    if group.leader_id != student.pk:
        messages.error(request, 'Only the group leader can edit this project.')
        return redirect('portal:gp_project')
    if is_gp_submission_locked(student.department):
        messages.error(request, 'The GP project submission deadline has passed. You can no longer edit your project.')
        return redirect('portal:gp_project')

    if request.method == 'POST':
        updated, errors = save_gp_submission(student, student.department, request.POST, group=group)
        if updated:
            messages.success(request, 'GP project updated successfully!')
            return redirect('portal:gp_project')
        for err in errors:
            messages.error(request, err)
        ctx = _gp_student_context(student, editing_group=group, post_data=request.POST)
        return render(request, 'portal/student/gp_project.html', ctx)

    ctx = _gp_student_context(student, editing_group=group)
    return render(request, 'portal/student/gp_project.html', ctx)


@login_required
def gp_group_delete(request, group_pk):
    if request.user.role != User.Role.STUDENT:
        return redirect('portal:dashboard')
    student = request.user.student_profile
    group = get_object_or_404(GPGroup, pk=group_pk)
    if group.leader_id != student.pk:
        messages.error(request, 'Only the group leader can delete this project.')
        return redirect('portal:gp_project')
    if is_gp_submission_locked(student.department):
        messages.error(request, 'The GP project submission deadline has passed. You can no longer delete your project.')
        return redirect('portal:gp_project')
    if request.method == 'POST':
        bundle = list(get_bundle_groups(group))
        for g in bundle:
            g.submissions.all().delete()
            g.delete()
        messages.success(request, 'GP project deleted.')
    return redirect('portal:gp_project')


@login_required
def gp_submit(request, group_pk):
    if request.user.role != User.Role.STUDENT:
        return redirect('portal:dashboard')
    student = request.user.student_profile
    group = get_object_or_404(GPGroup, pk=group_pk)
    if student not in group.members.all() and group.leader != student:
        messages.error(request, 'You are not a member of this group.')
        return redirect('portal:gp_project')

    template_id = request.GET.get('template') or request.POST.get('template_id')
    if not template_id:
        templates = FormTemplate.objects.filter(exam_type='GP', semester=group.department.semester)
        return render(request, 'portal/student/gp_select_template.html', {'group': group, 'templates': templates})

    template = get_object_or_404(FormTemplate, pk=template_id)
    existing = FormSubmission.objects.filter(group=group, template=template).first()
    initial = existing.field_data if existing else {}

    DynamicForm = build_dynamic_form(template, initial)
    if request.method == 'POST':
        form = DynamicForm(request.POST)
        if form.is_valid():
            data = {k: str(v) for k, v in form.cleaned_data.items()}
            if existing:
                existing.field_data = data
                existing.save()
            else:
                FormSubmission.objects.create(group=group, template=template, field_data=data)
            group.is_submitted = True
            group.save()
            messages.success(request, 'GP form submitted successfully!')
            return redirect('portal:gp_project')
    else:
        form = DynamicForm()

    return render(request, 'portal/student/gp_submit.html', {
        'form': form, 'group': group, 'template': template,
    })


# ─── Faculty: Mark Entry ───

@role_required(User.Role.FACULTY)
def mark_entry_list(request):
    faculty = request.user.faculty_profile
    duties = FacultyDutyAssignment.objects.filter(
        faculty=faculty, is_active=True,
    ).select_related('subject', 'department').order_by('-duty_date', 'batch')
    sessions = ExamSession.objects.filter(department=faculty.department, is_active=True)
    return render(request, 'portal/faculty/mark_entry_list.html', {
        'duties': duties,
        'sessions': sessions,
        'faculty': faculty,
    })


@role_required(User.Role.FACULTY)
def mark_entry_duty(request, duty_pk):
    duty = get_object_or_404(
        FacultyDutyAssignment,
        pk=duty_pk,
        faculty=request.user.faculty_profile,
        is_active=True,
        duty_role=FacultyDutyAssignment.DutyRole.INTERNAL,
    )
    students = sort_students_by_roll(Student.objects.filter(
        department=duty.department, batch=duty.batch,
    ))

    if request.method == 'POST':
        if not can_edit_duty_marks(request.user, duty):
            messages.error(request, 'Marks are verified and locked. You cannot edit them.')
            return redirect('portal:mark_entry_duty', duty_pk=duty_pk)

        action = request.POST.get('action', 'save')
        save_duty_marks(duty, request.POST, students, request.user)
        duty.refresh_from_db()

        if action == 'verify_lock':
            if duty.marks_locked:
                messages.info(request, 'Marks are already verified and locked.')
            elif get_duty_marks_status(duty) != 'completed':
                messages.warning(
                    request,
                    'Marks saved, but verify & lock requires all student entries to be complete.',
                )
            else:
                verify_and_lock_duty_marks(duty, request.user)
                messages.success(
                    request,
                    'Marks verified and locked. You can no longer edit them.',
                )
                return redirect('portal:mark_entry_duty', duty_pk=duty_pk)
        else:
            status = get_duty_marks_status(duty)
            if status == 'completed':
                messages.success(request, 'Marks saved successfully. Status: Completed.')
            else:
                messages.success(
                    request,
                    'Marks saved successfully. Status: Pending — complete all student entries.',
                )
        return redirect('portal:mark_entry_duty', duty_pk=duty_pk)

    return render(request, 'portal/faculty/mark_entry_duty.html', {
        **build_duty_marksheet_page_context(duty),
        'read_only': not can_edit_duty_marks(request.user, duty),
    })


def _internal_duties_in_scope(user, ctx, dept=None):
    qs = FacultyDutyAssignment.objects.filter(
        is_active=True,
        duty_role=FacultyDutyAssignment.DutyRole.INTERNAL,
    ).select_related('faculty', 'subject', 'department')
    if dept:
        qs = qs.filter(department=dept)
    elif user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        qs = qs.filter(department__semester=ctx['semester'])
    elif user.role == User.Role.DEPARTMENT_ADMIN and ctx.get('department'):
        qs = qs.filter(department=ctx['department'])
    return qs.order_by('-duty_date', 'batch', 'subject__name')


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def admin_marks_entry(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    exam_type = request.GET.get('exam_type', '').strip()
    subject_id = request.GET.get('subject', '').strip()

    duties_qs = FacultyDutyAssignment.objects.none()
    subjects = Subject.objects.none()
    if dept:
        duties_qs = _internal_duties_in_scope(request.user, ctx, dept)
        subjects = _subjects_for_dept_scope(dept, request.user, ctx)
        if exam_type in ('IPE', 'GP'):
            duties_qs = duties_qs.filter(exam_type=exam_type)
        if subject_id.isdigit():
            duties_qs = duties_qs.filter(subject_id=int(subject_id))

    duty_options = list(duties_qs)
    duty_rows = [
        {
            'duty': d,
            'marks_status': get_duty_marks_status(d),
            'can_edit': can_edit_duty_marks(request.user, d),
        }
        for d in duty_options
    ]

    return render(request, 'portal/admin/marks_entry.html', {
        'exam_type': exam_type,
        'selected_subject': subject_id,
        'subjects': subjects,
        'duty_options': duty_options,
        'duty_rows': duty_rows,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def admin_marks_entry_view(request, duty_pk):
    duty = get_object_or_404(
        FacultyDutyAssignment,
        pk=duty_pk,
        is_active=True,
        duty_role=FacultyDutyAssignment.DutyRole.INTERNAL,
    )
    ctx = get_user_context(request.user)
    if not _department_in_scope(request.user, ctx, duty.department):
        from django.http import Http404 as _Http404
        raise _Http404

    students = sort_students_by_roll(Student.objects.filter(
        department=duty.department, batch=duty.batch,
    ))
    if request.method == 'POST':
        if not can_edit_duty_marks(request.user, duty):
            messages.error(
                request,
                'Marks are verified and locked. Only semester admin or super admin can edit.',
            )
            return redirect('portal:admin_marks_entry_view', duty_pk=duty_pk)

        save_duty_marks(duty, request.POST, students, request.user)
        status = get_duty_marks_status(duty)
        if status == 'completed':
            messages.success(request, 'Marks saved successfully. Status: Completed.')
        else:
            messages.success(
                request,
                'Marks saved successfully. Status: Pending — complete all student entries.',
            )
        return redirect('portal:admin_marks_entry_view', duty_pk=duty_pk)

    can_edit = can_edit_duty_marks(request.user, duty)
    back_qs = f'?department={duty.department_id}&exam_type={duty.exam_type}&subject={duty.subject_id}'
    return render(request, 'portal/faculty/mark_entry_duty.html', {
        **build_duty_marksheet_page_context(duty),
        'read_only': not can_edit,
        'viewing_faculty': duty.faculty,
        'is_admin_entry': True,
        'back_url': reverse('portal:admin_marks_entry') + back_qs,
    })


@role_required(User.Role.FACULTY)
def mark_entry(request, session_pk):
    session = get_object_or_404(ExamSession, pk=session_pk)
    dept = request.user.faculty_profile.department
    batch = request.GET.get('batch', '')
    students = Student.objects.filter(department=dept)
    if batch:
        students = students.filter(batch=batch)
    batches = Student.objects.filter(department=dept).values_list('batch', flat=True).distinct()

    if request.method == 'POST':
        for stu in students:
            key = f'marks_{stu.pk}'
            if key in request.POST:
                marks_val = request.POST[key]
                remarks = request.POST.get(f'remarks_{stu.pk}', '')
                if marks_val:
                    MarkEntry.objects.update_or_create(
                        student=stu, exam_session=session,
                        defaults={
                            'marks_obtained': marks_val,
                            'remarks': remarks,
                            'entered_by': request.user,
                        }
                    )
        messages.success(request, 'Marks saved successfully.')
        return redirect('portal:mark_entry', session_pk=session_pk)

    existing_marks = {
        m.student_id: m
        for m in MarkEntry.objects.filter(exam_session=session, student__in=students)
    }
    student_rows = [
        {'student': s, 'mark': existing_marks.get(s.pk)}
        for s in students
    ]
    return render(request, 'portal/faculty/mark_entry.html', {
        'session': session, 'student_rows': student_rows,
        'batches': batches, 'batch_filter': batch,
    })


# ─── Exam Sessions ───

@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def exam_session_list(request):
    ctx = get_user_context(request.user)
    if request.user.role == User.Role.DEPARTMENT_ADMIN:
        dept = ctx['department']
        sessions = ExamSession.objects.filter(department=dept)
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
    else:
        dept = resolve_department(request.user, ctx, request)
        if dept:
            sessions = ExamSession.objects.filter(department=dept)
            subjects = Subject.objects.filter(
                Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
            )
        elif request.user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
            sessions = ExamSession.objects.filter(department__semester=ctx['semester'])
            subjects = Subject.objects.filter(semester=ctx['semester'])
        else:
            semester = resolve_semester(request.user, ctx, request)
            sessions = ExamSession.objects.filter(department__semester=semester) if semester else ExamSession.objects.all()
            subjects = Subject.objects.filter(semester=semester) if semester else Subject.objects.all()
    form = ExamSessionForm()
    form.fields['subject'].queryset = subjects
    return render(request, 'portal/exams/list.html', {
        'sessions': sessions,
        'form': form,
        **dept_filter_context(request.user, ctx, request, dept if request.user.role != User.Role.DEPARTMENT_ADMIN else ctx['department']),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def exam_session_create(request):
    ctx = get_user_context(request.user)
    if request.method == 'POST':
        form = ExamSessionForm(request.POST)
        dept = ctx.get('department')
        if request.user.role in [User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN]:
            dept_id = request.POST.get('department_id')
            if dept_id:
                dept = get_object_or_404(Department, pk=dept_id)
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        ) if dept else Subject.objects.all()
        form.fields['subject'].queryset = subjects
        if form.is_valid() and dept:
            exam = form.save(commit=False)
            exam.department = dept
            exam.save()
            messages.success(request, f'Exam session "{exam.name}" created.')
    return redirect('portal:exam_session_list')


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def exam_session_delete(request, pk):
    exam = get_object_or_404(ExamSession, pk=pk)
    if request.method == 'POST':
        exam.delete()
        messages.success(request, 'Exam session deleted.')
    return redirect('portal:exam_session_list')


def _subjects_for_dept_scope(dept, user, ctx):
    if dept:
        return Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
    if user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        return Subject.objects.filter(semester=ctx['semester'])
    semester = resolve_semester(user, ctx)
    return Subject.objects.filter(semester=semester) if semester else Subject.objects.all()


def _batches_for_dept_scope(dept, user, ctx):
    qs = Student.objects.all()
    if dept:
        qs = qs.filter(department=dept)
    elif user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        qs = qs.filter(department__semester=ctx['semester'])
    elif user.role == User.Role.DEPARTMENT_ADMIN and ctx.get('department'):
        qs = qs.filter(department=ctx['department'])
    return list(qs.values_list('batch', flat=True).distinct().order_by('batch'))


def _build_assignment_matrix(dept, subjects, batches):
    existing = {
        (a.batch, a.subject_id): a
        for a in FacultySubjectAssignment.objects.filter(
            department=dept, is_active=True,
        ).select_related('faculty', 'subject')
    }
    rows = []
    for batch in batches:
        subject_rows = []
        for subject in subjects:
            a = existing.get((batch, subject.pk))
            subject_rows.append({
                'subject': subject,
                'assignment': a,
                'faculty_id': a.faculty_id if a else '',
            })
        rows.append({'batch': batch, 'subjects': subject_rows})
    return rows


def _faculty_assignment_batches(faculty):
    return FacultySubjectAssignment.objects.filter(
        faculty=faculty, is_active=True,
    ).values_list('batch', flat=True).distinct().order_by('batch')


def _faculty_assignment_subjects(faculty, batch_filter=''):
    qs = FacultySubjectAssignment.objects.filter(faculty=faculty, is_active=True)
    if batch_filter:
        qs = qs.filter(batch=batch_filter)
    return Subject.objects.filter(
        pk__in=qs.values_list('subject_id', flat=True),
    ).distinct()


def _faculty_can_access_submission(faculty, subject_id, batch_filter):
    if not subject_id:
        return False
    qs = FacultySubjectAssignment.objects.filter(
        faculty=faculty, subject_id=subject_id, is_active=True,
    )
    if batch_filter:
        qs = qs.filter(batch=batch_filter)
    return qs.exists()


def _gp_groups_for_faculty(faculty, subject_id=None, batch_filter=None):
    assignments = FacultySubjectAssignment.objects.filter(
        faculty=faculty, is_active=True,
    )
    if batch_filter:
        assignments = assignments.filter(batch=batch_filter)
    if subject_id:
        assignments = assignments.filter(subject_id=subject_id)
    if not assignments.exists():
        return GPGroup.objects.none()

    q = Q()
    for a in assignments:
        q |= Q(
            department=a.department,
            subject=a.subject,
            leader__batch=a.batch,
            is_submitted=True,
        )
    return GPGroup.objects.filter(q).select_related(
        'subject', 'project_case', 'leader', 'department',
    ).prefetch_related('member_details__student', 'members').order_by(
        'leader__batch', '-updated_at',
    )


@role_required(User.Role.SEMESTER_ADMIN, User.Role.DEPARTMENT_ADMIN, User.Role.SUPER_ADMIN)
def assigned_subjects(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    subjects = _subjects_for_dept_scope(dept, request.user, ctx) if dept else Subject.objects.none()
    batches = _batches_for_dept_scope(dept, request.user, ctx) if dept else []
    faculty_qs = Faculty.objects.filter(department=dept).order_by('name') if dept else Faculty.objects.none()

    if request.method == 'POST' and request.POST.get('form_type') == 'save_assignments':
        if not dept:
            messages.error(request, 'Select a department first.')
        else:
            saved = 0
            for batch in batches:
                for subject in subjects:
                    key = f'faculty_{batch}_{subject.pk}'
                    faculty_id = request.POST.get(key, '').strip()
                    if faculty_id:
                        FacultySubjectAssignment.objects.update_or_create(
                            department=dept,
                            batch=batch,
                            subject=subject,
                            defaults={
                                'faculty_id': faculty_id,
                                'assigned_by': request.user,
                                'is_active': True,
                            },
                        )
                        saved += 1
                    else:
                        FacultySubjectAssignment.objects.filter(
                            department=dept, batch=batch, subject=subject,
                        ).update(is_active=False)
            messages.success(request, f'Subject faculty assignments saved ({saved} slots).')
            return redirect(f'{reverse("portal:assigned_subjects")}?department={dept.pk}')

    batch_rows = _build_assignment_matrix(dept, subjects, batches) if dept else []
    total_assignments = FacultySubjectAssignment.objects.filter(
        department=dept, is_active=True,
    ).count() if dept else 0

    return render(request, 'portal/subjects/assigned.html', {
        'batch_rows': batch_rows,
        'subjects': subjects,
        'batches': batches,
        'faculty_list': faculty_qs,
        'total_assignments': total_assignments,
        **dept_filter_context(request.user, ctx, request, dept),
    })


@role_required(User.Role.FACULTY)
def faculty_submissions(request):
    faculty = request.user.faculty_profile
    dept = faculty.department
    subject_id = request.GET.get('subject') or request.POST.get('subject')
    batch_filter = request.GET.get('batch', '')
    status_filter = request.GET.get('status', 'submitted')
    if status_filter not in ('submitted', 'pending'):
        status_filter = 'submitted'

    if request.GET.get('download') in ('1', 'attendance'):
        if not _faculty_can_access_submission(faculty, subject_id, batch_filter):
            messages.error(request, 'You are not assigned to this batch and subject.')
            return redirect('portal:faculty_submissions')
        groups = _gp_groups_for_faculty(faculty, subject_id, batch_filter)
        subject = Subject.objects.filter(pk=subject_id).first()
        if request.GET.get('download') == 'attendance':
            sem_label = dept.sheet_semester_label or (dept.semester.name if dept.semester else '')
            dept_label = dept.sheet_department_label or dept.name
            return generate_gp_attendance_workbook(
                dept, subject, sem_label, dept_label, batch_filter=batch_filter or None,
            )
        template = get_gp_template(dept)
        fname = f"gp_project_{subject.name.replace(' ', '_') if subject else 'all'}.xlsx"
        if batch_filter:
            fname = f"gp_project_{subject.name.replace(' ', '_')}_{batch_filter}.xlsx"
        return export_gp_submissions_excel(groups, template, fname, subject=subject)

    batches = list(_faculty_assignment_batches(faculty))
    subjects = _faculty_assignment_subjects(faculty, batch_filter)
    groups = _gp_groups_for_faculty(faculty, subject_id, batch_filter)
    pending_students = []
    if status_filter == 'pending' and subject_id and batch_filter:
        if _faculty_can_access_submission(faculty, subject_id, batch_filter):
            pending_students = list(get_pending_gp_students(dept, subject_id, batch_filter))
    my_assignments = FacultySubjectAssignment.objects.filter(
        faculty=faculty, is_active=True,
    ).select_related('subject').order_by('batch', 'subject__name')

    return render(request, 'portal/faculty/submissions.html', {
        'groups': groups,
        'pending_students': pending_students,
        'subjects': subjects,
        'batches': batches,
        'selected_subject': subject_id,
        'batch_filter': batch_filter,
        'status_filter': status_filter,
        'my_assignments': my_assignments,
        'department': dept,
        'faculty': faculty,
    })


def _gp_groups_in_scope(user, ctx, dept=None, subject_id=None, batch_filter=None):
    if dept:
        groups = GPGroup.objects.filter(department=dept, is_submitted=True)
    elif user.role == User.Role.SEMESTER_ADMIN:
        semester = ctx.get('semester')
        groups = GPGroup.objects.filter(department__semester=semester, is_submitted=True) if semester else GPGroup.objects.none()
    elif user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        groups = GPGroup.objects.filter(is_submitted=True)
    else:
        assigned = ctx.get('department')
        groups = GPGroup.objects.filter(department=assigned, is_submitted=True) if assigned else GPGroup.objects.none()

    if subject_id:
        groups = groups.filter(subject_id=subject_id)
    if batch_filter:
        groups = groups.filter(leader__batch=batch_filter)
    return groups.select_related(
        'subject', 'project_case', 'leader', 'department'
    ).prefetch_related('member_details__student', 'members').order_by('leader__batch', '-updated_at')


def _batches_for_submissions(user, ctx, dept):
    qs = Student.objects.all()
    if dept:
        qs = qs.filter(department=dept)
    elif user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        qs = qs.filter(department__semester=ctx['semester'])
    elif user.role == User.Role.DEPARTMENT_ADMIN and ctx.get('department'):
        qs = qs.filter(department=ctx['department'])
    return qs.values_list('batch', flat=True).distinct().order_by('batch')


def _pending_students_in_scope(user, ctx, dept=None, subject_id=None, batch_filter=None):
    """Students who have not submitted GP for batch + subject within admin scope."""
    if not subject_id or not batch_filter:
        return []

    departments = []
    if dept:
        departments = [dept]
    elif user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        departments = list(Department.objects.filter(semester=ctx['semester']))
    elif user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        semester = ctx.get('semester')
        if semester:
            departments = list(Department.objects.filter(semester=semester))
        else:
            departments = list(Department.objects.all())
    else:
        assigned = ctx.get('department')
        if assigned:
            departments = [assigned]

    pending = []
    for department in departments:
        pending.extend(get_pending_gp_students(department, subject_id, batch_filter))
    pending.sort(key=lambda s: (s.department.name, s.roll_no or '', s.name))
    return pending


@role_required(User.Role.DEPARTMENT_ADMIN, User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def view_submissions(request):
    ctx = get_user_context(request.user)
    dept = resolve_department(request.user, ctx, request)
    subject_id = request.GET.get('subject') or request.POST.get('subject')
    batch_filter = request.GET.get('batch', '')
    status_filter = request.GET.get('status', 'submitted')
    if status_filter not in ('submitted', 'pending'):
        status_filter = 'submitted'

    if request.GET.get('download') in ('1', 'attendance'):
        if not subject_id:
            messages.error(request, 'Please select a subject before downloading.')
            return redirect('portal:view_submissions')
        groups = _gp_groups_in_scope(request.user, ctx, dept, subject_id, batch_filter)
        subject = Subject.objects.filter(pk=subject_id).first()
        dept_obj = dept or (groups.first().department if groups.exists() else None)
        if request.GET.get('download') == 'attendance':
            if not dept_obj:
                messages.error(request, 'No submissions found for attendance sheet.')
                return redirect('portal:view_submissions')
            sem_label = dept_obj.sheet_semester_label or (dept_obj.semester.name if dept_obj.semester else '')
            dept_label = dept_obj.sheet_department_label or dept_obj.name
            return generate_gp_attendance_workbook(
                dept_obj, subject, sem_label, dept_label,
                batch_filter=batch_filter or None,
            )
        template = get_gp_template(dept_obj)
        fname = f"gp_project_{subject.name.replace(' ', '_') if subject else 'all'}.xlsx"
        if batch_filter:
            fname = f"gp_project_{subject.name.replace(' ', '_')}_{batch_filter}.xlsx"
        return export_gp_submissions_excel(groups, template, fname, subject=subject)

    if dept:
        subjects = Subject.objects.filter(
            Q(semester=dept.semester, department__isnull=True) | Q(department=dept)
        )
    elif request.user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        subjects = Subject.objects.filter(semester=ctx['semester'])
    else:
        semester = resolve_semester(request.user, ctx, request)
        subjects = Subject.objects.filter(semester=semester) if semester else Subject.objects.all()

    groups = _gp_groups_in_scope(request.user, ctx, dept, subject_id, batch_filter)
    batches = _batches_for_submissions(request.user, ctx, dept)
    pending_students = []
    if status_filter == 'pending':
        pending_students = _pending_students_in_scope(
            request.user, ctx, dept, subject_id, batch_filter,
        )

    return render(request, 'portal/department_admin/submissions.html', {
        'groups': groups,
        'pending_students': pending_students,
        'subjects': subjects,
        'batches': batches,
        'selected_subject': subject_id,
        'batch_filter': batch_filter,
        'status_filter': status_filter,
        **dept_filter_context(request.user, ctx, request, dept),
    })


# ─── External Examiner Registration ───

def become_external(request, form_pk=None):
    """Public page — external examiner registration."""
    if request.user.is_authenticated:
        return redirect('portal:dashboard')

    active_forms = ExternalRegistrationForm.objects.filter(
        is_active=True,
        semester__is_active=True,
    ).select_related('semester').order_by('-created_at')

    reg_form = None
    if form_pk:
        reg_form = active_forms.filter(pk=form_pk).first()
        if not reg_form:
            messages.error(request, 'This registration form is not available.')
            return redirect('portal:become_external')
    elif active_forms.count() == 1:
        reg_form = active_forms.first()

    if request.method == 'POST' and reg_form:
        FormClass = build_external_registration_form(reg_form)
        form = FormClass(request.POST, request.FILES)
        if form.is_valid():
            save_external_submission(reg_form, form.cleaned_data, request.FILES)
            messages.success(
                request,
                'Your external examiner registration has been submitted successfully. '
                'The institute will contact you shortly.',
            )
            return redirect('portal:become_external_success')
        fields = reg_form.fields.all()
    elif reg_form:
        FormClass = build_external_registration_form(reg_form)
        form = FormClass()
        fields = reg_form.fields.all()
    else:
        form = None
        fields = []

    return render(request, 'portal/external/public_form.html', {
        'reg_form': reg_form,
        'form': form,
        'fields': fields,
        'active_forms': active_forms,
    })


def become_external_success(request):
    if request.user.is_authenticated:
        return redirect('portal:dashboard')
    return render(request, 'portal/external/success.html')


@role_required(User.Role.SUPER_ADMIN)
def external_form_manage_list(request):
    forms_qs = ExternalRegistrationForm.objects.select_related('semester').prefetch_related('fields')
    semester = resolve_semester(request.user, get_user_context(request.user), request)
    if semester:
        forms_qs = forms_qs.filter(semester=semester)
    return render(request, 'portal/external/form_manage_list.html', {
        'registration_forms': forms_qs,
        'semester': semester,
    })


@role_required(User.Role.SUPER_ADMIN)
def external_form_create(request):
    if request.method == 'POST':
        form = ExternalRegistrationFormCreateForm(request.POST)
        if form.is_valid():
            reg_form = form.save(commit=False)
            reg_form.created_by = request.user
            reg_form.save()
            seed_external_form_fields(reg_form)
            messages.success(request, f'Form "{reg_form.title}" created with default fields.')
            return redirect('portal:external_form_fields', pk=reg_form.pk)
    else:
        form = ExternalRegistrationFormCreateForm()
    return render(request, 'portal/external/form_create.html', {'form': form})


@role_required(User.Role.SUPER_ADMIN)
def external_form_fields(request, pk):
    reg_form = get_object_or_404(ExternalRegistrationForm, pk=pk)

    if request.method == 'POST' and request.POST.get('form_type') == 'toggle_active':
        reg_form.is_active = not reg_form.is_active
        reg_form.save(update_fields=['is_active', 'updated_at'])
        status = 'activated' if reg_form.is_active else 'deactivated'
        messages.success(request, f'Form {status}.')
        return redirect('portal:external_form_fields', pk=pk)

    field_form = ExternalRegistrationFieldForm()
    if request.method == 'POST' and request.POST.get('form_type') == 'add_field':
        field_form = ExternalRegistrationFieldForm(request.POST)
        if field_form.is_valid():
            field = field_form.save(commit=False)
            field.form = reg_form
            field.field_name = unique_field_name(reg_form, field.field_label)
            field.save()
            messages.success(request, f'Field "{field.field_label}" added.')
            return redirect('portal:external_form_fields', pk=pk)

    if request.method == 'POST' and request.POST.get('form_type') == 'edit_field':
        fid = request.POST.get('field_id')
        field_obj = get_object_or_404(ExternalRegistrationField, pk=fid, form=reg_form)
        field_form = ExternalRegistrationFieldForm(request.POST, instance=field_obj)
        if field_form.is_valid():
            field_form.save()
            messages.success(request, 'Field updated.')
            return redirect('portal:external_form_fields', pk=pk)

    if request.method == 'POST' and request.POST.get('form_type') == 'delete_field':
        fid = request.POST.get('field_id')
        ExternalRegistrationField.objects.filter(pk=fid, form=reg_form).delete()
        messages.success(request, 'Field removed.')
        return redirect('portal:external_form_fields', pk=pk)

    if request.method == 'POST' and request.POST.get('form_type') == 'seed_defaults':
        count = seed_external_form_fields(reg_form)
        if count:
            messages.success(request, f'Added {count} default field(s).')
        else:
            messages.info(request, 'Default fields already exist.')
        return redirect('portal:external_form_fields', pk=pk)

    return render(request, 'portal/external/form_fields.html', {
        'reg_form': reg_form,
        'field_form': field_form,
        'form_fields': reg_form.fields.all(),
    })


@role_required(User.Role.SUPER_ADMIN)
def external_form_delete(request, pk):
    reg_form = get_object_or_404(ExternalRegistrationForm, pk=pk)
    if request.method == 'POST':
        title = reg_form.title
        reg_form.delete()
        messages.success(request, f'Form "{title}" deleted.')
    return redirect('portal:external_form_manage_list')


def _external_submissions_for_user(user, ctx, request=None):
    qs = ExternalRegistrationSubmission.objects.select_related('form', 'form__semester')
    if user.role == User.Role.SEMESTER_ADMIN and ctx.get('semester'):
        return qs.filter(form__semester=ctx['semester'])
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        semester = resolve_semester(user, ctx, request)
        if semester:
            qs = qs.filter(form__semester=semester)
        return qs
    return qs.none()


@role_required(User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def external_submission_list(request):
    ctx = get_user_context(request.user)
    submissions = _external_submissions_for_user(request.user, ctx, request)
    form_filter = request.GET.get('form')
    if form_filter:
        submissions = submissions.filter(form_id=form_filter)
    forms_qs = ExternalRegistrationForm.objects.all()
    if ctx.get('semester'):
        forms_qs = forms_qs.filter(semester=ctx['semester'])
    elif not (request.user.is_superuser or request.user.role == User.Role.SUPER_ADMIN):
        forms_qs = forms_qs.none()
    return render(request, 'portal/external/submissions_list.html', {
        'submissions': submissions,
        'registration_forms': forms_qs,
        'form_filter': form_filter,
        'semester': ctx.get('semester'),
    })


@role_required(User.Role.SEMESTER_ADMIN, User.Role.SUPER_ADMIN)
def external_submission_detail(request, pk):
    ctx = get_user_context(request.user)
    submission = get_object_or_404(
        ExternalRegistrationSubmission.objects.select_related('form', 'form__semester'),
        pk=pk,
    )
    allowed = _external_submissions_for_user(request.user, ctx, request).filter(pk=pk).exists()
    if not allowed:
        messages.error(request, 'You do not have access to this submission.')
        return redirect('portal:external_submission_list')
    fields = submission.form.fields.all()
    field_rows = []
    for field in fields:
        val = submission.field_data.get(field.field_name, '')
        field_rows.append({
            'field': field,
            'value': val or '—',
            'is_photo': field.field_type == ExternalRegistrationField.FieldType.PHOTO,
        })
    return render(request, 'portal/external/submission_detail.html', {
        'submission': submission,
        'field_rows': field_rows,
    })
