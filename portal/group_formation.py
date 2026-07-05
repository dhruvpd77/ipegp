"""GP group formation splits for attendance and marksheet generation."""
from .models import GPGroup, GPGroupFormation, Student, Subject
from .gp_utils import (
    _read_group_id,
    _group_id_value,
    _group_sort_key,
    build_gp_subject_selection_options,
    build_gp_download_subject_options,
    parse_subject_selection,
    subjects_for_department,
)


def formation_key_for_selection(department, subject_selection):
    """
    Map a (possibly individual) subject selection to the formation key it was saved under.

    Group Formations are saved with the combined practical value (e.g. P-12,13).
    Downloads may request a single practical subject (P-12); this returns the
    combined selection whose formation contains that subject so splits resolve.
    """
    if not department or not subject_selection:
        return subject_selection
    if get_formation(department, subject_selection):
        return subject_selection
    subject_ids = set(parse_subject_selection(subject_selection))
    if not subject_ids:
        return subject_selection
    for formation in GPGroupFormation.objects.filter(department=department):
        formation_ids = set(parse_subject_selection(formation.subject_selection))
        if subject_ids & formation_ids:
            return formation.subject_selection
    return subject_selection


def subject_selection_for_subject(department, subject):
    """Map a Subject instance to T-/P- selection value."""
    if not subject or not department:
        return ''
    for opt in build_gp_subject_selection_options(subjects_for_department(department)):
        if any(s['pk'] == subject.pk for s in opt['subjects']):
            return opt['value']
    if subject.subject_type == Subject.SubjectType.THEORY:
        return f'T-{subject.pk}'
    return ''


def resolve_subject_selection(department, subject=None, subject_selection=''):
    """Return subject_selection string from explicit value or Subject."""
    if subject_selection:
        return subject_selection.strip()
    if subject:
        return subject_selection_for_subject(department, subject)
    return ''


def _group_dedupe_key(group):
    if group.linked_batch_id:
        return str(group.linked_batch_id)
    gid = _read_group_id(group)
    if gid:
        return gid.lower()
    return f'pk-{group.pk}'


def get_unique_submitted_groups(department, subject_selection, batch=None):
    """Submitted GP groups for a subject selection, deduped for practical bundles."""
    subject_ids = parse_subject_selection(subject_selection)
    if not department or not subject_ids:
        return []

    groups = list(
        GPGroup.objects.filter(
            department=department,
            subject_id__in=subject_ids,
            is_submitted=True,
        ).select_related('leader', 'subject', 'project_case').prefetch_related(
            'member_details__student', 'members',
        )
    )
    if batch:
        groups = [g for g in groups if g.leader and g.leader.batch == batch]

    seen = {}
    unique = []
    for group in sorted(groups, key=_group_sort_key):
        key = _group_dedupe_key(group)
        if key in seen:
            continue
        seen[key] = True
        unique.append(group)
    return unique


def build_batch_group_id_lookup(department, subject_selection, batch):
    """
    Group ID map for a batch — same numbering as Project Details Excel export.
    Returns (entries, gid_to_group, group_to_gid).
    """
    groups = get_unique_submitted_groups(department, subject_selection, batch=batch)
    batch_counters = {}
    entries = []
    gid_to_group = {}
    group_to_gid = {}
    for group in sorted(groups, key=_group_sort_key):
        gid = _group_id_value(group, batch_counters)
        entry = {
            'group_id': gid,
            'group_pk': group.pk,
            'title': group.name,
            'member_count': group.member_details.count() or group.members.count(),
            'group': group,
        }
        entries.append(entry)
        gid_to_group[gid] = group
        gid_to_group[gid.lower()] = group
        group_to_gid[group.pk] = gid
        group_to_gid[group] = gid
    return entries, gid_to_group, group_to_gid


def list_batch_group_entries(department, subject_selection, batch):
    """Group IDs for a batch — sorted A1_1, A1_2, … A1_10 (matches Excel export)."""
    entries, _, _ = build_batch_group_id_lookup(department, subject_selection, batch)
    return [
        {k: v for k, v in e.items() if k != 'group'}
        for e in entries
    ]


def department_batches_for_formation(department):
    return list(
        Student.objects.filter(department=department)
        .values_list('batch', flat=True)
        .distinct()
        .order_by('batch')
    )


def get_formation(department, subject_selection):
    if not department or not subject_selection:
        return None
    return GPGroupFormation.objects.filter(
        department=department,
        subject_selection=subject_selection,
    ).first()


def save_batch_split(department, subject_selection, batch, split_count, splits, user=None):
    """Persist split configuration for one batch."""
    formation, _ = GPGroupFormation.objects.get_or_create(
        department=department,
        subject_selection=subject_selection,
        defaults={'batches_config': {}},
    )
    config = dict(formation.batches_config or {})
    config[batch] = {
        'split_count': int(split_count),
        'splits': splits,
    }
    formation.batches_config = config
    if user:
        formation.updated_by = user
    formation.save()
    return formation


def delete_batch_split(department, subject_selection, batch):
    formation = get_formation(department, subject_selection)
    if not formation:
        return
    config = dict(formation.batches_config or {})
    config.pop(batch, None)
    formation.batches_config = config
    formation.save()


def formation_summary_rows(department, subject_selection):
    """Table rows: batch, split label, comma-separated group IDs."""
    formation = get_formation(department, subject_selection)
    if not formation:
        return []
    rows = []
    for batch in sorted((formation.batches_config or {}).keys()):
        batch_cfg = formation.batches_config[batch]
        for split in batch_cfg.get('splits', []):
            gids = split.get('group_ids') or []
            rows.append({
                'batch': batch,
                'split_label': f"Split {split.get('index', '')}",
                'group_ids': gids,
                'group_ids_display': ', '.join(gids) if gids else '—',
            })
    return rows


def groups_for_split_ids(department, subject_selection, batch, group_ids):
    """Resolve GPGroup objects for saved group IDs (preserves split order)."""
    if not group_ids:
        return []
    _, gid_to_group, _ = build_batch_group_id_lookup(department, subject_selection, batch)
    resolved = []
    for gid in group_ids:
        group = gid_to_group.get(gid) or gid_to_group.get(str(gid).lower())
        if group:
            resolved.append(group)
    return resolved


def formation_sheet_title(group_ids):
    """Excel tab name: comma-separated group IDs in the split (max 31 chars)."""
    if not group_ids:
        return 'Empty'
    title = ', '.join(str(gid) for gid in group_ids)
    if len(title) <= 31:
        return title
    # Keep as many full group IDs as fit within Excel's 31-character limit.
    parts = []
    length = 0
    for gid in group_ids:
        part = str(gid)
        extra = len(part) if not parts else len(part) + 2
        if length + extra > 28:
            break
        parts.append(part)
        length += extra
    if parts:
        return ', '.join(parts) + '...'
    return str(group_ids[0])[:31]


def iter_formation_sheets(department, subject_selection, batch_filter=None,
                          group_subject_selection=None):
    """
    Yield (sheet_title, batch, split_index, groups, group_to_gid) for each saved split.
    Sheet order: batches alphabetically, then split index.

    `subject_selection` locates the saved formation (splits/group IDs).
    `group_subject_selection` (optional) selects which subject's GPGroup rows to
    resolve — used for individual-subject downloads (FSD-II vs FCSP-II) that share
    the same combined formation but have per-subject project details.
    """
    formation = get_formation(department, subject_selection)
    if not formation or not formation.batches_config:
        return

    group_selection = group_subject_selection or subject_selection

    batches = sorted(formation.batches_config.keys())
    if batch_filter:
        batches = [b for b in batches if b == batch_filter]

    for batch in batches:
        _, _, group_to_gid = build_batch_group_id_lookup(department, group_selection, batch)
        batch_cfg = formation.batches_config.get(batch) or {}
        splits = batch_cfg.get('splits') or []
        for split in sorted(splits, key=lambda s: s.get('index', 0)):
            idx = split.get('index', 1)
            group_ids = split.get('group_ids') or []
            groups = groups_for_split_ids(department, group_selection, batch, group_ids)
            if not groups:
                continue
            title = formation_sheet_title(group_ids)
            yield title, batch, idx, groups, group_to_gid


def students_for_split_groups(groups):
    """Unique students from groups in roll order."""
    seen = set()
    students = []
    for group in groups:
        details = list(group.member_details.select_related('student').order_by('student__roll_no'))
        if not details:
            for m in group.members.order_by('roll_no'):
                details.append(type('Detail', (), {'student': m})())
        for detail in details:
            stu = detail.student
            if stu.pk in seen:
                continue
            seen.add(stu.pk)
            students.append(stu)
    return sorted(students, key=lambda s: (s.roll_no or '', s.enrollment_no or ''))


def primary_subject_for_selection(department, subject_selection):
    subject_ids = parse_subject_selection(subject_selection)
    if not subject_ids:
        return None
    return Subject.objects.filter(pk=subject_ids[0]).first()


def list_formation_split_options(department, subject_selection):
    """Dropdown options for GP duty: batch (group1, group2, …) per saved split."""
    formation = get_formation(department, subject_selection)
    if not formation or not formation.batches_config:
        return []
    options = []
    for batch in sorted(formation.batches_config.keys()):
        batch_cfg = formation.batches_config.get(batch) or {}
        for split in sorted(batch_cfg.get('splits') or [], key=lambda s: s.get('index', 0)):
            gids = split.get('group_ids') or []
            if not gids:
                continue
            idx = split.get('index', 1)
            label = f"{batch} ({', '.join(gids)})"
            options.append({
                'value': f'{batch}|{idx}',
                'label': label,
                'batch': batch,
                'split_index': idx,
                'group_ids': gids,
            })
    return options


def parse_split_option_value(value):
    """Parse 'A1|2' into (batch, split_index)."""
    if not value or '|' not in value:
        return None, None
    batch, idx = value.split('|', 1)
    try:
        return batch.strip(), int(idx)
    except (TypeError, ValueError):
        return None, None
