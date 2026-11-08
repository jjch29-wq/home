"""Per-joint quantities used by the central branch's daily entry form."""


def quantity_per_joint(method, pipe, size_lengths):
    pipe = str(pipe).strip().upper()
    if pipe.isdigit():
        pipe += 'A'
    if pipe not in size_lengths:
        return None
    if method == 'RT':
        # Same rule as daily_work_log_tab: RT up to 150A uses three films.
        size = int(pipe.removesuffix('A'))
        return 3 if size <= 150 else None
    return size_lengths[pipe]
