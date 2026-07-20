"""Email external examiner login credentials (IPE / GP duty) for marks entry."""
from __future__ import annotations

from email.mime.image import MIMEImage
from pathlib import Path

from django.conf import settings
from django.core.mail import EmailMultiAlternatives
from django.utils.html import escape

from .models import GeneratedCredential, User


FROM_EMAIL = getattr(
    settings,
    'DEFAULT_FROM_EMAIL',
    'LJ Institute of Engineering & Technology <ipeljiet@gmail.com>',
)
REPLY_TO = getattr(settings, 'EMAIL_REPLY_TO', ['ipeljiet@gmail.com'])

# Live portal (always use this in external examiner emails)
PORTAL_BASE_URL = getattr(
    settings,
    'PORTAL_PUBLIC_BASE_URL',
    'https://ljietgp.pythonanywhere.com',
).rstrip('/')
PORTAL_LOGIN_URL = f'{PORTAL_BASE_URL}/login/faculty/'
PORTAL_MARKS_URL = f'{PORTAL_BASE_URL}/marks/'


def portal_urls():
    return PORTAL_LOGIN_URL, PORTAL_MARKS_URL


def _email_logo_path(filename):
    return Path(settings.BASE_DIR) / 'static' / 'images' / 'email' / filename


def get_faculty_plain_password(faculty):
    """Return stored plain password for faculty login, or None."""
    if not faculty or not faculty.user_id:
        return None
    cred = (
        GeneratedCredential.objects.filter(user_id=faculty.user_id)
        .order_by('-generated_at')
        .first()
    )
    return (cred.plain_password if cred else None) or None


def normalize_external_username(email):
    """External examiner username = their email (lowercased)."""
    return (email or '').strip().lower()


def ensure_external_username_is_email(faculty, email=None):
    """
    Keep faculty.user.username equal to their email.
    Returns (username, error_message_or_None).
    """
    email_addr = normalize_external_username(email or faculty.email)
    if not email_addr or '@' not in email_addr:
        return None, f'Valid email is required for {faculty.name} (used as username).'
    if not faculty.user_id:
        return None, f'No login generated yet for {faculty.name}. Generate ID/Password first.'

    user = faculty.user
    if user.username != email_addr:
        conflict = User.objects.filter(username__iexact=email_addr).exclude(pk=user.pk).exists()
        if conflict:
            return None, f'Username/email {email_addr} is already used by another account.'
        user.username = email_addr
        user.email = email_addr
        user.save(update_fields=['username', 'email'])
    elif not user.email:
        user.email = email_addr
        user.save(update_fields=['email'])

    if faculty.email != email_addr:
        faculty.email = email_addr
        faculty.save(update_fields=['email'])
    return email_addr, None


def build_credentials_email_html(*, name, username, password, login_url, marks_url, exam_label='IPE/GP'):
    safe_name = escape(name or 'Sir/Madam')
    safe_user = escape(username or '')
    safe_pass = escape(password or '')
    safe_login = escape(login_url or '')
    safe_marks = escape(marks_url or '')
    safe_exam = escape(exam_label or 'IPE/GP')
    safe_site = escape(PORTAL_BASE_URL)

    return f"""\
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>IPE/GP Portal Login</title>
</head>
<body style="margin:0;padding:0;background:#eef2f7;font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1f2937;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#eef2f7;padding:28px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="max-width:640px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 10px 28px rgba(15,23,42,0.10);">
          <tr>
            <td style="background:#0b3d5c;padding:22px 28px;">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="vertical-align:middle;width:54%;">
                    <img src="cid:lju_logo" alt="LJ University" style="display:block;max-width:260px;width:100%;height:auto;background:#ffffff;border-radius:8px;padding:8px 10px;">
                  </td>
                  <td style="vertical-align:middle;text-align:right;width:46%;">
                    <img src="cid:ljiet_logo" alt="LJIET" style="display:inline-block;max-height:78px;width:auto;background:#ffffff;border-radius:8px;padding:6px;">
                  </td>
                </tr>
              </table>
              <div style="margin-top:16px;color:#ffffff;">
                <div style="font-size:12px;letter-spacing:0.08em;text-transform:uppercase;opacity:0.85;">IPE / GP Portal</div>
                <div style="font-size:20px;font-weight:700;margin-top:4px;line-height:1.3;">
                  External Examiner Login Credentials
                </div>
                <div style="font-size:13px;margin-top:6px;opacity:0.92;">{safe_exam} Marks Entry Access</div>
              </div>
            </td>
          </tr>
          <tr>
            <td style="padding:28px 32px 10px 32px;">
              <p style="margin:0 0 14px 0;font-size:16px;line-height:1.6;">
                Dear <strong>{safe_name}</strong>,
              </p>
              <p style="margin:0 0 16px 0;font-size:15px;line-height:1.7;color:#374151;">
                Greetings from <strong>L. J. Institute of Engineering &amp; Technology</strong>
                (LJ University). Your portal login is ready for
                <strong>{safe_exam}</strong> marks entry.
              </p>

              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin:8px 0 18px 0;border:1px solid #dbe3ee;border-radius:10px;overflow:hidden;">
                <tr>
                  <td colspan="2" style="background:#f3f7fb;padding:12px 16px;font-weight:700;font-size:14px;color:#0b3d5c;border-bottom:1px solid #dbe3ee;">
                    Login Details
                  </td>
                </tr>
                <tr>
                  <td style="padding:12px 16px;width:34%;font-size:13px;color:#6b7280;border-bottom:1px solid #eef2f7;">Website</td>
                  <td style="padding:12px 16px;font-size:14px;font-weight:600;border-bottom:1px solid #eef2f7;">
                    <a href="{safe_site}" style="color:#145a86;text-decoration:none;">{safe_site}</a>
                  </td>
                </tr>
                <tr>
                  <td style="padding:12px 16px;font-size:13px;color:#6b7280;border-bottom:1px solid #eef2f7;">Username</td>
                  <td style="padding:12px 16px;font-size:15px;font-weight:700;border-bottom:1px solid #eef2f7;font-family:Consolas,Monaco,monospace;color:#0b3d5c;">{safe_user}</td>
                </tr>
                <tr>
                  <td style="padding:12px 16px;font-size:13px;color:#6b7280;">Password</td>
                  <td style="padding:12px 16px;font-size:18px;font-weight:700;letter-spacing:0.12em;font-family:Consolas,Monaco,monospace;color:#b91c1c;">{safe_pass}</td>
                </tr>
              </table>

              <p style="margin:0 0 10px 0;font-size:13px;color:#6b7280;">
                Your <strong>Username</strong> is your registered email address.
              </p>

              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin:6px 0 18px 0;">
                <tr>
                  <td align="center" style="padding:6px;">
                    <a href="{safe_login}" style="display:inline-block;background:#145a86;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 20px;border-radius:8px;">
                      Open Faculty Login
                    </a>
                  </td>
                </tr>
                <tr>
                  <td align="center" style="padding:6px;">
                    <a href="{safe_marks}" style="display:inline-block;background:#0f766e;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 20px;border-radius:8px;">
                      Go to Marks Entry
                    </a>
                  </td>
                </tr>
              </table>

              <p style="margin:0 0 8px 0;font-size:13px;line-height:1.7;color:#4b5563;">
                <strong>Login:</strong>
                <a href="{safe_login}" style="color:#145a86;word-break:break-all;">{safe_login}</a>
              </p>
              <p style="margin:0 0 16px 0;font-size:13px;line-height:1.7;color:#4b5563;">
                <strong>Marks Entry:</strong>
                <a href="{safe_marks}" style="color:#145a86;word-break:break-all;">{safe_marks}</a>
              </p>

              <p style="margin:0 0 8px 0;font-size:14px;line-height:1.7;color:#374151;">
                After login, open <strong>My Duties</strong> to enter marks for your assigned batch / split.
                Please keep these credentials confidential.
              </p>

              <p style="margin:22px 0 0 0;font-size:15px;line-height:1.6;">
                Warm regards,<br>
                <strong>IPE Coordination Team</strong><br>
                L. J. Institute of Engineering &amp; Technology<br>
                <a href="mailto:ipeljiet@gmail.com" style="color:#145a86;text-decoration:none;">ipeljiet@gmail.com</a>
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:16px 32px 22px 32px;background:#f8fafc;border-top:1px solid #e5e7eb;font-size:12px;color:#6b7280;line-height:1.6;">
              Portal: <a href="{safe_site}" style="color:#145a86;text-decoration:none;">{safe_site}</a><br>
              This email was sent from the IPE/GP Portal of L. J. Institute of Engineering &amp; Technology.
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"""


def build_credentials_email_text(*, name, username, password, login_url, marks_url, exam_label='IPE/GP'):
    return (
        f'Dear {name or "Sir/Madam"},\n\n'
        f'Greetings from L. J. Institute of Engineering & Technology (LJ University).\n'
        f'Your {exam_label} Portal login credentials for marks entry are below.\n\n'
        f'Website: {PORTAL_BASE_URL}\n'
        f'Username (your email): {username}\n'
        f'Password: {password}\n\n'
        f'Faculty Login: {login_url}\n'
        f'Marks Entry: {marks_url}\n\n'
        f'After login, open My Duties to fill marks.\n\n'
        f'Warm regards,\n'
        f'IPE Coordination Team\n'
        f'L. J. Institute of Engineering & Technology\n'
        f'ipeljiet@gmail.com\n'
    )


def _attach_logo(email_msg, filename, cid):
    path = _email_logo_path(filename)
    if not path.exists():
        return False
    with open(path, 'rb') as fh:
        img = MIMEImage(fh.read())
    img.add_header('Content-ID', f'<{cid}>')
    img.add_header('Content-Disposition', 'inline', filename=filename)
    email_msg.attach(img)
    return True


def send_external_credentials_email(
    faculty,
    *,
    login_url=None,
    marks_url=None,
    exam_label='IPE/GP',
    to_email=None,
):
    """
    Send username/password + portal links to one external faculty.
    Username is always the examiner's email.
    Returns (ok: bool, message: str).
    """
    email_addr = normalize_external_username(to_email or faculty.email)
    if not email_addr or '@' not in email_addr:
        return False, f'No valid email for {faculty.name}. Add an email first (used as username).'

    if not faculty.user_id:
        return False, f'No login generated yet for {faculty.name}. Generate ID/Password first.'

    username, sync_err = ensure_external_username_is_email(faculty, email_addr)
    if sync_err:
        return False, sync_err

    password = get_faculty_plain_password(faculty)
    if not password:
        return False, f'Password not found for {faculty.name}. Generate login again.'

    if not getattr(settings, 'EMAIL_HOST_PASSWORD', ''):
        return (
            False,
            'Email is not configured. Set DJANGO_EMAIL_HOST_PASSWORD for ipeljiet@gmail.com.',
        )

    login_url = login_url or PORTAL_LOGIN_URL
    marks_url = marks_url or PORTAL_MARKS_URL
    mail_subject = (
        f'{exam_label} Portal Login — Marks Entry | '
        f'LJ Institute of Engineering and Technology'
    )
    context = dict(
        name=faculty.name,
        username=username,
        password=password,
        login_url=login_url,
        marks_url=marks_url,
        exam_label=exam_label,
    )
    email = EmailMultiAlternatives(
        subject=mail_subject,
        body=build_credentials_email_text(**context),
        from_email=FROM_EMAIL,
        to=[email_addr],
        reply_to=REPLY_TO,
    )
    email.attach_alternative(build_credentials_email_html(**context), 'text/html')
    email.mixed_subtype = 'related'
    _attach_logo(email, 'lju-logo.png', 'lju_logo')
    _attach_logo(email, 'ljiet-logo-sif.png', 'ljiet_logo')
    email.send(fail_silently=False)
    return True, f'Login credentials emailed to {faculty.name} ({email_addr}).'
