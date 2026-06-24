from django.contrib import admin
from django.contrib.auth.admin import UserAdmin

from .models import (
    User, Semester, SemesterAdminAssignment, Department, DepartmentAdminAssignment,
    Subject, Faculty, Student, FormTemplate, FormField, GPGroup, GPGroupMemberDetail,
    FormSubmission, ProjectCase,
    ExamSession, MarkEntry, GeneratedCredential, AttendanceSheetTemplate,
    MarksheetTemplate, FacultyDutyAssignment, DutyScheduleUpload,
    SubjectSyllabus, SubjectPaper, FacultySubjectAssignment,
    ExternalRegistrationForm, ExternalRegistrationField, ExternalRegistrationSubmission,
)


@admin.register(User)
class CustomUserAdmin(UserAdmin):
    list_display = ['username', 'email', 'role', 'is_active']
    list_filter = ['role', 'is_active']
    fieldsets = UserAdmin.fieldsets + (('Role', {'fields': ('role', 'phone')}),)
    add_fieldsets = UserAdmin.add_fieldsets + (('Role', {'fields': ('role', 'phone')}),)


admin.site.register(Semester)
admin.site.register(SemesterAdminAssignment)
admin.site.register(Department)
admin.site.register(DepartmentAdminAssignment)
admin.site.register(Subject)
admin.site.register(Faculty)
admin.site.register(Student)
admin.site.register(FormTemplate)
admin.site.register(FormField)
admin.site.register(GPGroup)
admin.site.register(GPGroupMemberDetail)
admin.site.register(ProjectCase)
admin.site.register(FormSubmission)
admin.site.register(ExamSession)
admin.site.register(MarkEntry)
admin.site.register(GeneratedCredential)
admin.site.register(AttendanceSheetTemplate)
admin.site.register(MarksheetTemplate)
admin.site.register(FacultyDutyAssignment)
admin.site.register(DutyScheduleUpload)
admin.site.register(SubjectSyllabus)
admin.site.register(SubjectPaper)
admin.site.register(FacultySubjectAssignment)
admin.site.register(ExternalRegistrationForm)
admin.site.register(ExternalRegistrationField)
admin.site.register(ExternalRegistrationSubmission)