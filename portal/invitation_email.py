"""Send IPE invitation emails with invitation PDF + subject syllabus attachments."""
from __future__ import annotations

import mimetypes
import os

from django.conf import settings
from django.core.mail import EmailMultiAlternatives
from django.db.models import Q
from django.utils.html import escape

from .invitation_letter import build_invitation_pdf
from .models import SubjectSyllabus


FROM_EMAIL = getattr(
    settings,
    'DEFAULT_FROM_EMAIL',
    'LJ Institute of Engineering & Technology <ipeljiet@gmail.com>',
)
REPLY_TO = getattr(settings, 'EMAIL_REPLY_TO', ['ipeljiet@gmail.com'])


def resolve_subject_syllabus(subject, department=None):
    """Prefer department-specific active syllabus, then semester-wide."""
    if not subject:
        return None
    qs = SubjectSyllabus.objects.filter(
        subject=subject,
        is_active=True,
        semester=subject.semester,
    ).order_by('-created_at')
    if department:
        dept_one = qs.filter(department=department).first()
        if dept_one:
            return dept_one
        return qs.filter(department__isnull=True).first()
    return qs.filter(Q(department__isnull=True) | Q(department__isnull=False)).first()


def _safe_attachment_name(name, fallback='attachment'):
    safe = ''.join(ch if ch.isalnum() or ch in (' ', '_', '-', '.') else '_' for ch in (name or fallback))
    safe = ' '.join(safe.split()).strip() or fallback
    return safe[:180]


def build_invitation_email_html(batch, faculty):
    """Professional HTML body for the invitation email."""
    name = escape(faculty.name or 'Sir/Madam')
    subject_name = escape(batch.subject_name or (batch.subject.name if batch.subject_id else 'IPE'))
    practical = escape(batch.practical_date or '—')
    exam_time = escape(batch.exam_time or '—')
    branch = escape(batch.branch or '—')
    letter_date = escape(batch.letter_date or '—')
    designation = escape(faculty.designation or '')
    college = escape(faculty.college_name or '')

    return f"""\
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>IPE Invitation</title>
</head>
<body style="margin:0;padding:0;background:#f0f2f5;font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1f2937;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f0f2f5;padding:24px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="max-width:640px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 8px 24px rgba(15,23,42,0.08);">
          <tr>
            <td style="background:linear-gradient(135deg,#0b3d5c 0%,#145a86 100%);padding:28px 32px;color:#ffffff;">
              <div style="font-size:13px;letter-spacing:0.08em;text-transform:uppercase;opacity:0.85;">L. J. University</div>
              <div style="font-size:22px;font-weight:700;margin-top:6px;line-height:1.3;">
                L. J. Institute of Engineering &amp; Technology
              </div>
              <div style="font-size:14px;margin-top:10px;opacity:0.92;">
                Internal Practical Examination — External Examiner Invitation
              </div>
            </td>
          </tr>
          <tr>
            <td style="padding:28px 32px 8px 32px;">
              <p style="margin:0 0 16px 0;font-size:16px;line-height:1.6;">
                Dear <strong>{name}</strong>{', ' + designation if designation else ''},
              </p>
              <p style="margin:0 0 16px 0;font-size:15px;line-height:1.7;color:#374151;">
                Greetings from <strong>L. J. Institute of Engineering &amp; Technology</strong>.
                We are pleased to invite you as an <strong>External Examiner</strong> for the
                Internal Practical Examination.
              </p>
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin:20px 0;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">
                <tr>
                  <td colspan="2" style="background:#f8fafc;padding:12px 16px;font-weight:700;font-size:14px;color:#0b3d5c;border-bottom:1px solid #e5e7eb;">
                    Examination Details
                  </td>
                </tr>
                <tr>
                  <td style="padding:10px 16px;width:38%;font-size:13px;color:#6b7280;border-bottom:1px solid #f3f4f6;">Subject</td>
                  <td style="padding:10px 16px;font-size:14px;font-weight:600;border-bottom:1px solid #f3f4f6;">{subject_name}</td>
                </tr>
                <tr>
                  <td style="padding:10px 16px;font-size:13px;color:#6b7280;border-bottom:1px solid #f3f4f6;">Date of Practical</td>
                  <td style="padding:10px 16px;font-size:14px;font-weight:600;border-bottom:1px solid #f3f4f6;">{practical}</td>
                </tr>
                <tr>
                  <td style="padding:10px 16px;font-size:13px;color:#6b7280;border-bottom:1px solid #f3f4f6;">Exam Time</td>
                  <td style="padding:10px 16px;font-size:14px;font-weight:600;border-bottom:1px solid #f3f4f6;">{exam_time}</td>
                </tr>
                <tr>
                  <td style="padding:10px 16px;font-size:13px;color:#6b7280;border-bottom:1px solid #f3f4f6;">Branch</td>
                  <td style="padding:10px 16px;font-size:14px;font-weight:600;border-bottom:1px solid #f3f4f6;">{branch}</td>
                </tr>
                <tr>
                  <td style="padding:10px 16px;font-size:13px;color:#6b7280;">Letter Date</td>
                  <td style="padding:10px 16px;font-size:14px;font-weight:600;">{letter_date}</td>
                </tr>
              </table>
              {"<p style='margin:0 0 12px 0;font-size:14px;color:#4b5563;'><strong>College:</strong> " + college + "</p>" if college else ""}
              <p style="margin:0 0 16px 0;font-size:15px;line-height:1.7;color:#374151;">
                Please find attached:
              </p>
              <ul style="margin:0 0 20px 20px;padding:0;font-size:14px;line-height:1.8;color:#374151;">
                <li>Your official <strong>Invitation Letter</strong> (PDF)</li>
                <li>The <strong>Syllabus</strong> of the subject (if available)</li>
              </ul>
              <p style="margin:0 0 8px 0;font-size:15px;line-height:1.7;color:#374151;">
                We look forward to your kind presence and valuable contribution.
              </p>
              <p style="margin:24px 0 0 0;font-size:15px;line-height:1.6;">
                Warm regards,<br>
                <strong>IPE Coordination Team</strong><br>
                L. J. Institute of Engineering &amp; Technology<br>
                <a href="mailto:ipeljiet@gmail.com" style="color:#145a86;text-decoration:none;">ipeljiet@gmail.com</a>
              </p>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 32px 24px 32px;background:#f8fafc;border-top:1px solid #e5e7eb;font-size:12px;color:#6b7280;line-height:1.6;">
              This email was sent from the IPE/GP Portal of L. J. Institute of Engineering &amp; Technology.
              If you received this message in error, please ignore it or reply to this address.
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"""


def build_invitation_email_text(batch, faculty):
    name = faculty.name or 'Sir/Madam'
    subject_name = batch.subject_name or (batch.subject.name if batch.subject_id else 'IPE')
    return (
        f'Dear {name},\n\n'
        f'Greetings from L. J. Institute of Engineering & Technology.\n'
        f'You are invited as an External Examiner for the Internal Practical Examination.\n\n'
        f'Subject: {subject_name}\n'
        f'Date of Practical: {batch.practical_date}\n'
        f'Exam Time: {batch.exam_time}\n'
        f'Branch: {batch.branch}\n\n'
        f'Please find attached the Invitation Letter and Subject Syllabus.\n\n'
        f'Warm regards,\n'
        f'IPE Coordination Team\n'
        f'L. J. Institute of Engineering & Technology\n'
        f'ipeljiet@gmail.com\n'
    )


def send_invitation_email(batch, faculty, syllabus=None):
    """
    Send invitation email to one external examiner.
    Returns (ok: bool, message: str, syllabus_attached: bool).
    """
    to_email = (faculty.email or '').strip()
    if not to_email:
        return False, 'No email address for this examiner.', False

    if not getattr(settings, 'EMAIL_HOST_PASSWORD', ''):
        return (
            False,
            'Email is not configured. Set DJANGO_EMAIL_HOST_PASSWORD (Gmail App Password) for ipeljiet@gmail.com.',
            False,
        )

    subject_name = batch.subject_name or (batch.subject.name if batch.subject_id else 'IPE')
    mail_subject = (
        f'Invitation Letter — External Examiner for IPE ({subject_name}) '
        f'| LJ Institute of Engineering and Technology'
    )

    email = EmailMultiAlternatives(
        subject=mail_subject,
        body=build_invitation_email_text(batch, faculty),
        from_email=FROM_EMAIL,
        to=[to_email],
        reply_to=REPLY_TO,
    )
    email.attach_alternative(build_invitation_email_html(batch, faculty), 'text/html')

    pdf_bytes = build_invitation_pdf(batch, faculty)
    email.attach(
        _safe_attachment_name(faculty.invite_filename, 'INVITE_LETTER.pdf'),
        pdf_bytes,
        'application/pdf',
    )

    syllabus_attached = False
    if syllabus and syllabus.file:
        try:
            syllabus.file.open('rb')
            file_bytes = syllabus.file.read()
            syllabus.file.close()
        except Exception:
            file_bytes = None
        if file_bytes:
            base = os.path.basename(syllabus.file.name) or f'{subject_name}_Syllabus'
            filename = _safe_attachment_name(base, f'{subject_name}_Syllabus')
            content_type = mimetypes.guess_type(filename)[0] or 'application/octet-stream'
            email.attach(filename, file_bytes, content_type)
            syllabus_attached = True

    email.send(fail_silently=False)
    if syllabus_attached:
        return True, f'Invitation email sent to {to_email} (letter + syllabus attached).', True
    return True, f'Invitation email sent to {to_email} (letter attached; no syllabus found for this subject).', False
