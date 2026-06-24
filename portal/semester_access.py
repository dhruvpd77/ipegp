"""Semester active/inactive access rules for portal users."""
from .models import User


def get_user_semester(user):
    """Return the Semester linked to this user's profile, if any."""
    if user.role == User.Role.SEMESTER_ADMIN:
        try:
            return user.semester_admin_assignment.semester
        except Exception:
            return None
    if user.role == User.Role.DEPARTMENT_ADMIN:
        try:
            return user.department_admin_assignment.department.semester
        except Exception:
            return None
    if user.role == User.Role.FACULTY:
        try:
            return user.faculty_profile.department.semester
        except Exception:
            return None
    if user.role == User.Role.STUDENT:
        try:
            return user.student_profile.department.semester
        except Exception:
            return None
    return None


def is_portal_user_blocked_by_inactive_semester(user):
    """
    Faculty, students (and external examiners logged in as faculty) cannot
    use the portal when their semester is inactive.
    """
    if not user.is_authenticated:
        return False
    if user.is_superuser or user.role in (
        User.Role.SUPER_ADMIN,
        User.Role.SEMESTER_ADMIN,
        User.Role.DEPARTMENT_ADMIN,
    ):
        return False
    if user.role not in (User.Role.FACULTY, User.Role.STUDENT):
        return False
    semester = get_user_semester(user)
    if semester is None:
        return False
    return not semester.is_active


def can_manage_semester_status(user, semester):
    """Super admin or assigned semester admin may toggle semester status."""
    if user.is_superuser or user.role == User.Role.SUPER_ADMIN:
        return True
    if user.role == User.Role.SEMESTER_ADMIN:
        try:
            return user.semester_admin_assignment.semester_id == semester.pk
        except Exception:
            return False
    return False
