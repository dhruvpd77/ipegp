from django.urls import path
from . import views

app_name = 'portal'

urlpatterns = [
    path('', views.home, name='home'),
    path('become-external/', views.become_external, name='become_external'),
    path('become-external/success/', views.become_external_success, name='become_external_success'),
    path('become-external/<int:form_pk>/', views.become_external, name='become_external_form'),
    path('login/', views.login_view, name='login'),
    path('login/<str:role>/', views.login_view, name='login_role'),
    path('change-password/', views.change_password, name='change_password'),
    path('semester-inactive/', views.semester_inactive, name='semester_inactive'),
    path('logout/', views.logout_view, name='logout'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('user-management/', views.user_management, name='user_management'),

    # Super Admin
    path('semesters/', views.semester_list, name='semester_list'),
    path('semesters/create/', views.semester_create, name='semester_create'),
    path('semesters/<int:pk>/edit/', views.semester_edit, name='semester_edit'),
    path('semesters/<int:pk>/toggle-active/', views.semester_toggle_active, name='semester_toggle_active'),
    path('semesters/<int:pk>/delete/', views.semester_delete, name='semester_delete'),
    path('semesters/<int:semester_pk>/create-admin/', views.create_semester_admin, name='create_semester_admin'),

    # Semester Admin
    path('departments/', views.department_list, name='department_list'),
    path('departments/create/', views.department_create, name='department_create'),
    path('departments/<int:pk>/delete/', views.department_delete, name='department_delete'),
    path('departments/<int:dept_pk>/create-admin/', views.create_department_admin, name='create_department_admin'),

    # Subjects
    path('subjects/assigned/', views.assigned_subjects, name='assigned_subjects'),
    path('subjects/', views.subject_list, name='subject_list'),
    path('subjects/create/', views.subject_create, name='subject_create'),
    path('subjects/<int:pk>/edit/', views.subject_edit, name='subject_edit'),
    path('subjects/<int:pk>/delete/', views.subject_delete, name='subject_delete'),

    # Students & Faculty upload
    path('upload/students/', views.upload_students, name='upload_students'),
    path('upload/faculty/', views.upload_faculty, name='upload_faculty'),
    # Faculty portal (must be before faculty/<int:pk>/ admin routes)
    path('faculty/syllabus/', views.faculty_syllabus, name='faculty_syllabus'),
    path('faculty/syllabus/<int:pk>/view/', views.faculty_syllabus_view, name='faculty_syllabus_view'),
    path('faculty/syllabus/<int:pk>/file/', views.faculty_syllabus_file, name='faculty_syllabus_file'),
    path('faculty/papers/', views.faculty_papers, name='faculty_papers'),
    path('faculty/papers/<int:pk>/view/', views.faculty_paper_view, name='faculty_paper_view'),
    path('faculty/papers/<int:pk>/file/', views.faculty_paper_file, name='faculty_paper_file'),
    path('faculty/', views.faculty_list, name='faculty_list'),
    path('faculty/create/', views.faculty_create, name='faculty_create'),
    path('faculty/generate-credentials/', views.generate_faculty_credentials, name='generate_faculty_credentials'),
    path('faculty/<int:pk>/edit/', views.faculty_edit, name='faculty_edit'),
    path('faculty/<int:pk>/delete/', views.faculty_delete, name='faculty_delete'),
    path('students/', views.student_list, name='student_list'),
    path('students/create/', views.student_create, name='student_create'),
    path('students/<int:pk>/edit/', views.student_edit, name='student_edit'),
    path('students/<int:pk>/delete/', views.student_delete, name='student_delete'),
    path('students/generate-credentials/', views.generate_student_credentials, name='generate_student_credentials'),
    path('students/export/', views.export_students, name='export_students'),

    # Attendance Sheets
    path('attendance/', views.attendance_sheets, name='attendance_sheets'),
    path('attendance/template/<int:pk>/delete/', views.attendance_template_delete, name='attendance_template_delete'),
    path('group-formations/', views.group_formations, name='group_formations'),

    # Marksheet Templates
    path('marksheets/', views.marksheet_templates, name='marksheet_templates'),
    path('marksheets/template/<int:pk>/delete/', views.marksheet_template_delete, name='marksheet_template_delete'),
    path('marksheets/final-download/', views.final_marksheet_download, name='final_marksheet_download'),

    # Faculty Duty Assignment
    path('duty/', views.faculty_duty_list, name='faculty_duty_list'),
    path('duty/<int:pk>/delete/', views.faculty_duty_delete, name='faculty_duty_delete'),
    path('duty-gp/', views.faculty_duty_gp, name='faculty_duty_gp'),
    path('duty-gp/<int:pk>/delete/', views.faculty_duty_gp_delete, name='faculty_duty_gp_delete'),

    # Syllabus & Papers
    path('syllabus-papers/', views.syllabus_papers_manage, name='syllabus_papers_manage'),
    path('syllabus/<int:pk>/delete/', views.syllabus_delete, name='syllabus_delete'),
    path('papers/<int:pk>/delete/', views.paper_delete, name='paper_delete'),

    # Form Templates
    path('forms/', views.form_template_list, name='form_template_list'),
    path('forms/create/', views.form_template_create, name='form_template_create'),
    path('forms/<int:pk>/fields/', views.form_template_fields, name='form_template_fields'),
    path('forms/<int:pk>/delete/', views.form_template_delete, name='form_template_delete'),
    path('forms/cases/', views.project_case_list, name='project_case_list'),
    path('forms/gp-deadline/', views.gp_deadline_settings, name='gp_deadline_settings'),

    # Student portal
    path('profile/', views.student_profile, name='student_profile'),
    path('gp/', views.gp_project, name='gp_project'),
    path('gp/<int:group_pk>/edit/', views.gp_group_edit, name='gp_group_edit'),
    path('gp/<int:group_pk>/delete/', views.gp_group_delete, name='gp_group_delete'),
    path('gp/<int:group_pk>/submit/', views.gp_submit, name='gp_submit'),

    # Faculty submissions (scoped to assigned batches)
    path('faculty/submissions/', views.faculty_submissions, name='faculty_submissions'),

    # Faculty mark entry
    path('marks/', views.mark_entry_list, name='mark_entry_list'),
    path('marks/duty/<int:duty_pk>/', views.mark_entry_duty, name='mark_entry_duty'),
    path('marks/<int:session_pk>/', views.mark_entry, name='mark_entry'),

    # Admin marks entry (view faculty marksheets) — not under /admin/ (Django admin conflict)
    path('marks-entry/', views.admin_marks_entry, name='admin_marks_entry'),
    path('marks-entry/duty/<int:duty_pk>/', views.admin_marks_entry_view, name='admin_marks_entry_view'),

    # Exams
    path('exams/', views.exam_session_list, name='exam_session_list'),
    path('exams/create/', views.exam_session_create, name='exam_session_create'),
    path('exams/<int:pk>/delete/', views.exam_session_delete, name='exam_session_delete'),

    # Submissions
    path('submissions/', views.view_submissions, name='view_submissions'),

    # GP Analytics
    path('analytics/', views.gp_analytics, name='gp_analytics'),

    # External examiner registration
    path('external-forms/', views.external_form_manage_list, name='external_form_manage_list'),
    path('external-forms/create/', views.external_form_create, name='external_form_create'),
    path('external-forms/<int:pk>/fields/', views.external_form_fields, name='external_form_fields'),
    path('external-forms/<int:pk>/delete/', views.external_form_delete, name='external_form_delete'),
    path('external-submissions/', views.external_submission_list, name='external_submission_list'),
    path('external-submissions/<int:pk>/', views.external_submission_detail, name='external_submission_detail'),
]
