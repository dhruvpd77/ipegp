"""IPE external examiner invitation letter — Excel import + PDF generation.

PDF layout matches ANIRUDDHSINH KARSHANBHAI DODIYA_INVITE LETTER.docx /
blank.docx: official LJ letterhead header+footer, centered Sub / schedule
block, numbered terms, right-aligned date and signature.
"""
import io
import re
from pathlib import Path

import openpyxl
from django.conf import settings
from django.http import HttpResponse
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Frame, Image, KeepTogether, Paragraph,
    BaseDocTemplate, PageTemplate, Spacer,
)
from xml.sax.saxutils import escape as xml_escape

def _p(text):
    """Escape user/static text for ReportLab Paragraph markup."""
    return xml_escape(str(text or ''))



def _norm_header(value):
    text = str(value or '').strip().lower()
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return ' '.join(text.split())


HEADER_ALIASES = {
    'name': {
        'name', 'faculty name', 'faculty', 'external name', 'examiner name',
        'external faculty name', 'full name',
    },
    'designation': {'designation', 'designation name', 'title', 'post'},
    'college_name': {
        'college', 'college name', 'institute', 'institute name',
        'organization', 'organisation', 'college / institute',
    },
    'city_state': {
        'city state', 'city / state', 'city/state', 'city', 'state',
        'city and state', 'address', 'location',
    },
    'email': {'email', 'email id', 'e mail', 'mail'},
}


def _map_columns(headers):
    mapping = {}
    for idx, raw in enumerate(headers):
        key = _norm_header(raw)
        if not key:
            continue
        for field, aliases in HEADER_ALIASES.items():
            if key in aliases and field not in mapping:
                mapping[field] = idx
                break
    return mapping


def parse_invitation_excel(uploaded_file):
    """
    Read Excel rows into list of dicts:
    name, designation, college_name, city_state, email.
    """
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError('Excel file is empty.')

    header_row_idx = 0
    for i, row in enumerate(rows[:5]):
        cells = [str(c or '') for c in row]
        if any('name' in _norm_header(c) for c in cells):
            header_row_idx = i
            break

    headers = rows[header_row_idx]
    mapping = _map_columns(headers)
    if 'name' not in mapping:
        raise ValueError(
            'Excel must include a Name column (e.g. Name / Faculty Name / External Name).'
        )

    faculties = []
    for row in rows[header_row_idx + 1:]:
        if not row or all(c is None or str(c).strip() == '' for c in row):
            continue

        def cell(field, default=''):
            idx = mapping.get(field)
            if idx is None or idx >= len(row):
                return default
            val = row[idx]
            return str(val).strip() if val is not None else default

        name = cell('name')
        if not name:
            continue
        faculties.append({
            'name': name,
            'designation': cell('designation'),
            'college_name': cell('college_name'),
            'city_state': cell('city_state'),
            'email': cell('email'),
        })

    if not faculties:
        raise ValueError('No faculty rows found in the Excel file.')
    return faculties


def _salutation_name(name):
    text = (name or '').strip()
    if not text:
        return 'Sir/Madam'
    upper = text.upper()
    if upper.startswith(('MR.', 'MR ', 'MS.', 'MS ', 'MRS.', 'MRS ', 'DR.', 'DR ', 'PROF.')):
        return text
    return f'Mr. {text}'


def _asset(*parts):
    return Path(settings.BASE_DIR) / 'static' / 'images' / 'invite_letter' / Path(*parts)


def _styles():
    font = 'Times-Roman'
    bold = 'Times-Bold'
    return {
        'date': ParagraphStyle(
            'InviteDate', fontName=bold, fontSize=11, leading=13,
            alignment=TA_RIGHT, spaceAfter=6,
        ),
        'to_line': ParagraphStyle(
            'InviteTo', fontName=bold, fontSize=11, leading=13,
            alignment=TA_LEFT, spaceAfter=0,
        ),
        'name': ParagraphStyle(
            'InviteName', fontName=bold, fontSize=11, leading=13,
            alignment=TA_LEFT, spaceAfter=0,
        ),
        'sub': ParagraphStyle(
            'InviteSub', fontName=bold, fontSize=11, leading=13,
            alignment=TA_CENTER, spaceBefore=8, spaceAfter=6,
        ),
        'intro': ParagraphStyle(
            'InviteIntro', fontName=font, fontSize=10.5, leading=13,
            alignment=TA_LEFT, spaceAfter=6,
        ),
        'detail': ParagraphStyle(
            'InviteDetail', fontName=bold, fontSize=11, leading=13,
            alignment=TA_CENTER, spaceAfter=1,
        ),
        'term': ParagraphStyle(
            'InviteTerm', fontName=font, fontSize=10, leading=12.5,
            alignment=TA_JUSTIFY, spaceAfter=2, leftIndent=4, firstLineIndent=0,
        ),
        'term_bold': ParagraphStyle(
            'InviteTermBold', fontName=bold, fontSize=10, leading=12.5,
            alignment=TA_JUSTIFY, spaceAfter=2, leftIndent=22,
        ),
        'method': ParagraphStyle(
            'InviteMethod', fontName=font, fontSize=10, leading=12.5,
            alignment=TA_JUSTIFY, spaceAfter=1, leftIndent=18,
        ),
        'method_hdr': ParagraphStyle(
            'InviteMethodHdr', fontName=bold, fontSize=10, leading=12.5,
            alignment=TA_JUSTIFY, spaceAfter=2, leftIndent=4,
        ),
        'sign': ParagraphStyle(
            'InviteSign', fontName=bold, fontSize=11, leading=13,
            alignment=TA_RIGHT, spaceAfter=0,
        ),
        'sign_for': ParagraphStyle(
            'InviteFor', fontName=bold, fontSize=11, leading=13,
            alignment=TA_RIGHT, spaceBefore=6, spaceAfter=1,
        ),
    }


# Terms match the DOCX numbered list (1–5 + Method as 6, then a–g).
STATIC_TERMS = [
    (
        'You are entitled for Examination Remuneration and Transport Allowance '
        'as per prevailing Institute / Trust norms from exam center.'
    ),
    'Examiners out of Gujarat will be provided with TO & FRO Air Fare and Lodging.',
    (
        'In case of any extreme/unavoidable circumstances of not able to perform duty, '
        'kindly inform at least prior ONE day to the Subject In-charge Faculty.'
    ),
    (
        'We have also attached our LJU BE Syllabus, Exams Question Papers, Practical List, '
        'and Practical Marksheet Format for your reference. Examiners are free to modify '
        'the Practical List as per the syllabus or Marksheet format and mail us the updated '
        'practical list at least before TWO days of Exam start date.'
    ),
    'One LJ Internal Examiner will assist you throughout the Exam duration.',
]

REMUNERATION_LINE = (
    'Remuneration for Conducting Practical Exam: Rs. 1000 per Batch'
)

METHOD_ITEMS = [
    'Students will be called Roll No wise and asked to pick up practical chits randomly.',
    (
        'They are allowed to change the Practical title chit only once in first 15 mins. '
        'For this chance, 2 Marks will be deducted from secured practical marks. After '
        'changing the chit, they strictly have to perform the latest practical.'
    ),
    'If students report late (10 mins after scheduled time), then ONE mark will be deducted.',
    (
        'Examiner has to compulsory conduct practical examination and can ask questions '
        'related to the practical performed by students. If required brief viva of a student '
        'can be taken related to the practical/subject.'
    ),
    (
        'Examiner can ask students for minor updates to check whether his/her practical '
        'concept is clear or not.'
    ),
    (
        'Submit duly signed hand-written Marksheet, Attendance report in a sealed envelope '
        'to Internal Examiner, before leaving exam center.'
    ),
    'Also Prepare Marksheet Soft Copy as per the format available with Internal Examiner.',
]


def _draw_letterhead(canvas, doc):
    """Paint official blank.docx letterhead (header + footer) as full-page background."""
    bg = _asset('page_bg_clean.jpg')
    if not bg.exists():
        bg = _asset('page_bg.jpg')
    if bg.exists():
        canvas.saveState()
        canvas.drawImage(
            str(bg), 0, 0, width=A4[0], height=A4[1],
            preserveAspectRatio=False, mask='auto',
        )
        canvas.restoreState()


def build_invitation_pdf(batch, faculty):
    """Return PDF bytes for one faculty invitation letter (DOCX layout)."""
    buffer = io.BytesIO()
    page_w, page_h = A4

    # Content frame sits in the white band between letterhead header & footer.
    # Blank letterhead crop: header ~235/1755, footer from ~1638/1755 of page image.
    top_margin = 38 * mm
    bottom_margin = 28 * mm
    left_margin = 18 * mm
    right_margin = 18 * mm

    frame = Frame(
        left_margin,
        bottom_margin,
        page_w - left_margin - right_margin,
        page_h - top_margin - bottom_margin,
        id='body',
        showBoundary=0,
    )
    doc = BaseDocTemplate(
        buffer,
        pagesize=A4,
        title=faculty.invite_filename,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )
    doc.addPageTemplates([
        PageTemplate(id='letter', frames=[frame], onPage=_draw_letterhead),
    ])

    styles = _styles()
    story = []

    story.append(Paragraph(f'Date: {_p(batch.letter_date)}', styles['date']))
    story.append(Paragraph('To,', styles['to_line']))
    story.append(Paragraph(_p(_salutation_name(faculty.name)), styles['name']))
    if faculty.designation:
        story.append(Paragraph(_p(faculty.designation.rstrip(',') + ','), styles['to_line']))
    if faculty.college_name:
        story.append(Paragraph(_p(faculty.college_name.rstrip(',') + ','), styles['to_line']))
    if faculty.city_state:
        story.append(Paragraph(_p(faculty.city_state), styles['to_line']))
    if faculty.email:
        story.append(Paragraph(f'Email: {_p(faculty.email)}', styles['to_line']))

    story.append(Paragraph(f'Sub: {_p(batch.subject_line)}', styles['sub']))
    story.append(Paragraph('Dear Sir/Madam,', styles['intro']))
    story.append(Paragraph(
        'I am glad to appoint you as an external examiner to conduct practical examination '
        'of B.E Sem- III as per the following schedule:',
        styles['intro'],
    ))

    story.append(Paragraph(f'Subject Name: {_p(batch.subject_name)}', styles['detail']))
    story.append(Paragraph(f'Date of Practical: {_p(batch.practical_date)}', styles['detail']))
    story.append(Paragraph(f'Branch: {_p(batch.branch)}', styles['detail']))
    story.append(Paragraph(f'Exam Time: {_p(batch.exam_time)}', styles['detail']))
    story.append(Spacer(1, 3 * mm))

    # Numbered terms matching DOCX (1–5, remuneration under 1, then 6 + a–g)
    story.append(Paragraph(f'<b>1.</b>  {_p(STATIC_TERMS[0])}', styles['term']))
    story.append(Paragraph(_p(REMUNERATION_LINE), styles['term_bold']))
    for i, text in enumerate(STATIC_TERMS[1:], start=2):
        story.append(Paragraph(f'<b>{i}.</b>  {_p(text)}', styles['term']))
    story.append(Paragraph('<b>6.</b>  Method of conducting practical exam:', styles['method_hdr']))

    for letter, text in zip('abcdefg', METHOD_ITEMS):
        story.append(Paragraph(f'<b>{letter}.</b>  {_p(text)}', styles['method']))

    # Signature block — right aligned; uses admin-saved signature when available
    sig_path, advisor_title, advisor_name = _resolve_signature(batch)
    sign_bits = [Paragraph('for', styles['sign_for'])]
    if sig_path:
        try:
            sign_bits.append(Image(str(sig_path), width=22 * mm, height=16 * mm, hAlign='RIGHT'))
        except Exception:
            pass
    sign_bits.append(Paragraph(_p(advisor_title), styles['sign']))
    sign_bits.append(Paragraph(_p(advisor_name), styles['sign']))
    story.append(KeepTogether(sign_bits))

    doc.build(story)
    return buffer.getvalue()


def _resolve_signature(batch=None):
    """Return (path_or_None, advisor_title, advisor_name) for the PDF signature block."""
    title = 'Advisor'
    name = '(Mr Rohit Patel)'
    try:
        from .models import IPEInvitationSignature
        sig = None
        if batch is not None and getattr(batch, 'subject_id', None):
            sig = IPEInvitationSignature.objects.filter(subject_id=batch.subject_id).first()
        if sig is None:
            # Legacy global row (subject NULL) if present
            sig = IPEInvitationSignature.objects.filter(subject__isnull=True).order_by('pk').first()
        if sig:
            title = sig.advisor_title or title
            name = sig.advisor_name or name
            if sig.signature:
                path = Path(sig.signature.path)
                if path.exists():
                    return path, title, name
    except Exception:
        pass
    fallback = _asset('signature.jpg')
    if fallback.exists():
        return fallback, title, name
    return None, title, name


def invitation_pdf_response(batch, faculty):
    pdf_bytes = build_invitation_pdf(batch, faculty)
    response = HttpResponse(pdf_bytes, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{faculty.invite_filename}"'
    return response


def _draw_thanks_letterhead(canvas, doc):
    """Paint LJU_Letterhead.docx background for thank-you letters."""
    bg = _asset('thanks_page_bg_clean.jpg')
    if not bg.exists():
        bg = _asset('lh_image1.jpeg')
    if bg.exists():
        canvas.saveState()
        canvas.drawImage(
            str(bg), 0, 0, width=A4[0], height=A4[1],
            preserveAspectRatio=False, mask='auto',
        )
        canvas.restoreState()


def _thanks_styles():
    font = 'Times-Roman'
    bold = 'Times-Bold'
    return {
        'date': ParagraphStyle(
            'ThanksDate', fontName=bold, fontSize=12, leading=16,
            alignment=TA_RIGHT, spaceAfter=14,
        ),
        'to_line': ParagraphStyle(
            'ThanksTo', fontName=font, fontSize=12, leading=15,
            alignment=TA_LEFT, spaceAfter=0,
        ),
        'name': ParagraphStyle(
            'ThanksName', fontName=bold, fontSize=12, leading=15,
            alignment=TA_LEFT, spaceAfter=0,
        ),
        'conducted': ParagraphStyle(
            'ThanksConducted', fontName=bold, fontSize=12, leading=16,
            alignment=TA_LEFT, spaceBefore=14, spaceAfter=10,
        ),
        'dear': ParagraphStyle(
            'ThanksDear', fontName=font, fontSize=12, leading=16,
            alignment=TA_LEFT, spaceAfter=8,
        ),
        'body': ParagraphStyle(
            'ThanksBody', fontName=font, fontSize=12, leading=17,
            alignment=TA_JUSTIFY, spaceAfter=10,
        ),
        'closing': ParagraphStyle(
            'ThanksClosing', fontName=font, fontSize=12, leading=16,
            alignment=TA_LEFT, spaceBefore=4, spaceAfter=4,
        ),
        'sign': ParagraphStyle(
            'ThanksSign', fontName=bold, fontSize=12, leading=15,
            alignment=TA_LEFT, spaceAfter=0,
        ),
        'sign_label': ParagraphStyle(
            'ThanksSignLabel', fontName=font, fontSize=12, leading=15,
            alignment=TA_LEFT, spaceBefore=18, spaceAfter=2,
        ),
    }


THANKS_PARA_1 = (
    'LJIET is a forward-looking reputed engineering institution with an objective to provide '
    'excellence in higher education across the state of Gujarat.'
)
THANKS_PARA_2 = (
    'We are highly thankful that you accepted our invitation and spared time from your busy '
    'schedule to come to our institute for evaluating Project/Practical Examination. It was '
    'indeed a very enriching experience for our students as well as faculties.'
)
THANKS_PARA_3 = (
    'We look forward to continue our relationship and request you to give benefit of your '
    'expertise to our students from time to time.'
)


def build_thanks_pdf(batch, faculty):
    """Return PDF bytes for one faculty thank-you letter (thank letter.docx + LJU letterhead)."""
    from reportlab.platypus.flowables import _listWrapOn

    buffer = io.BytesIO()
    page_w, page_h = A4

    # LJU letterhead header is taller (blue curve); leave space for footer block.
    top_margin = 42 * mm
    bottom_margin = 32 * mm
    left_margin = 20 * mm
    right_margin = 20 * mm
    frame_w = page_w - left_margin - right_margin
    frame_h = page_h - top_margin - bottom_margin

    frame = Frame(
        left_margin,
        bottom_margin,
        frame_w,
        frame_h,
        id='thanks_body',
        showBoundary=0,
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )
    doc = BaseDocTemplate(
        buffer,
        pagesize=A4,
        title=faculty.thanks_filename,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )
    doc.addPageTemplates([
        PageTemplate(id='thanks', frames=[frame], onPage=_draw_thanks_letterhead),
    ])

    styles = _thanks_styles()
    body = []

    # Date follows Date of Practical (per product requirement)
    letter_date = batch.practical_date or batch.letter_date
    body.append(Paragraph(f'Date: {_p(letter_date)}', styles['date']))
    body.append(Paragraph('To,', styles['to_line']))
    body.append(Paragraph(_p(_salutation_name(faculty.name)), styles['name']))
    if faculty.designation:
        body.append(Paragraph(_p(faculty.designation.rstrip(',') + ','), styles['to_line']))
    if faculty.college_name:
        body.append(Paragraph(_p(faculty.college_name.rstrip(',') + ','), styles['to_line']))
    if faculty.city_state:
        body.append(Paragraph(_p(faculty.city_state), styles['to_line']))

    body.append(Paragraph(
        f'Project/Practical Exam Conducted on: {_p(batch.practical_date)}',
        styles['conducted'],
    ))
    body.append(Paragraph('Dear Sir/Madam,', styles['dear']))
    body.append(Paragraph(_p(THANKS_PARA_1), styles['body']))
    body.append(Paragraph(_p(THANKS_PARA_2), styles['body']))
    body.append(Paragraph(_p(THANKS_PARA_3), styles['body']))
    body.append(Paragraph('Thanking you,', styles['closing']))
    body.append(Paragraph('Yours truly,', styles['closing']))

    sig_path, advisor_title, advisor_name = _resolve_signature(batch)
    if sig_path:
        try:
            body.append(Spacer(1, 2 * mm))
            body.append(Image(str(sig_path), width=28 * mm, height=18 * mm, hAlign='LEFT'))
        except Exception:
            pass
    body.append(Paragraph(_p(advisor_title or 'Advisor'), styles['sign']))
    body.append(Paragraph(_p(advisor_name or '(Mr. Rohit Patel)'), styles['sign']))

    # Measure then pad top so the block sits in the vertical middle of the page
    _w, block_h = _listWrapOn(body, frame_w, None)
    top_pad = max(0, (frame_h - block_h) / 2.0)
    doc.build([Spacer(1, top_pad), KeepTogether(body)])
    return buffer.getvalue()


def thanks_pdf_response(batch, faculty):
    pdf_bytes = build_thanks_pdf(batch, faculty)
    response = HttpResponse(pdf_bytes, content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{faculty.thanks_filename}"'
    return response
