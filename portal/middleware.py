from django.shortcuts import redirect

from .semester_access import is_portal_user_blocked_by_inactive_semester


class SemesterAccessMiddleware:
    """Block faculty/students when their semester is marked inactive."""

    ALLOWED_URL_NAMES = frozenset({
        'logout',
        'change_password',
        'semester_inactive',
    })

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated and is_portal_user_blocked_by_inactive_semester(request.user):
            match = getattr(request, 'resolver_match', None)
            url_name = match.url_name if match else None
            if url_name not in self.ALLOWED_URL_NAMES:
                return redirect('portal:semester_inactive')
        return self.get_response(request)
