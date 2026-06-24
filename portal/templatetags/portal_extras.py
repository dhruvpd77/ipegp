from django import template

register = template.Library()

USER_MANAGEMENT_URLS = frozenset({
    'user_management',
    'student_list', 'student_edit', 'student_create',
    'faculty_list', 'faculty_edit', 'faculty_create',
    'subject_list', 'subject_create', 'subject_edit',
    'assigned_subjects',
    'syllabus_papers_manage', 'syllabus_delete', 'paper_delete',
    'project_case_list',
})


@register.filter
def is_academic_mgmt_active(url_name):
    return url_name in USER_MANAGEMENT_URLS


@register.simple_tag(takes_context=True)
def is_user_management_active(context, url_name=None):
    """True when current page is the hub or any user-management sub-page."""
    if url_name:
        return url_name in USER_MANAGEMENT_URLS
    request = context.get('request')
    if not request or not getattr(request, 'resolver_match', None):
        return False
    return request.resolver_match.url_name in USER_MANAGEMENT_URLS
