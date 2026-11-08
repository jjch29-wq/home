"""Selection flags leave the daily crew roster and its input values intact."""


def worker_is_selected(group):
    flag = getattr(group, 'participation_var', None)
    return bool(flag.get()) if flag is not None else True


def select_workers(app, selected):
    for index in range(1, 11):
        group = getattr(app, f'worker_group{index}', None)
        if group is not None and hasattr(group, 'participation_var'):
            group.participation_var.set(bool(selected))
