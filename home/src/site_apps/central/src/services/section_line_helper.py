"""Helper functions for cascading company -> section -> line options in Central branch app."""
import json
import os


def extract_company_section_line_mapping(history_data):
    """
    Extract cascading relationships from daily_work_history data.

    Returns:
        dict with keys:
            'company_to_sections': dict[company, list[str]]
            'company_section_to_lines': dict[(company, section), list[str]]
            'section_to_lines': dict[section, list[str]]
            'all_companies': list[str]
            'all_sections': list[str]
            'all_lines': list[str]
    """
    company_to_sections = {}
    company_section_to_lines = {}
    section_to_lines = {}
    method_to_lines = {}
    method_to_sections = {}
    company_method_to_sections = {}
    company_method_to_lines = {}
    section_method_to_lines = {}
    company_section_method_to_lines = {}

    all_companies = set()
    all_sections = set()
    all_lines = set()
    all_methods = set()

    for date_str, day_data in (history_data or {}).items():
        if not isinstance(day_data, dict):
            continue
        for row in day_data.get('ndt_results', []):
            if not isinstance(row, dict):
                continue
            company = str(row.get('업체', '') or '').strip()
            section = str(row.get('구간', '') or '').strip()
            line = str(row.get('라인번호', '') or '').strip()
            method = str(row.get('검사방법', '') or row.get('방법', '') or row.get('ndt_type', '') or '').strip().upper()

            if company:
                all_companies.add(company)
            if section:
                all_sections.add(section)
            if line:
                all_lines.add(line)
            if method:
                all_methods.add(method)

            if company and section:
                company_to_sections.setdefault(company, set()).add(section)
            if section and line:
                section_to_lines.setdefault(section, set()).add(line)
            if company and section and line:
                company_section_to_lines.setdefault((company, section), set()).add(line)
            elif company and line:
                company_section_to_lines.setdefault((company, ''), set()).add(line)

            if method and line:
                method_to_lines.setdefault(method, set()).add(line)
            if method and section:
                method_to_sections.setdefault(method, set()).add(section)
            if company and method and section:
                company_method_to_sections.setdefault((company, method), set()).add(section)
            if company and method and line:
                company_method_to_lines.setdefault((company, method), set()).add(line)
            if section and method and line:
                section_method_to_lines.setdefault((section, method), set()).add(line)
            if company and section and method and line:
                company_section_method_to_lines.setdefault((company, section, method), set()).add(line)

    c_to_s = {c: sorted(s) for c, s in company_to_sections.items()}
    cs_to_l = {k: sorted(v) for k, v in company_section_to_lines.items()}
    s_to_l = {s: sorted(lines) for s, lines in section_to_lines.items()}
    m_to_l = {m: sorted(lines) for m, lines in method_to_lines.items()}
    m_to_s = {m: sorted(secs) for m, secs in method_to_sections.items()}
    cm_to_s = {k: sorted(v) for k, v in company_method_to_sections.items()}
    cm_to_l = {k: sorted(v) for k, v in company_method_to_lines.items()}
    sm_to_l = {k: sorted(v) for k, v in section_method_to_lines.items()}
    csm_to_l = {k: sorted(v) for k, v in company_section_method_to_lines.items()}

    return {
        'company_to_sections': c_to_s,
        'company_section_to_lines': cs_to_l,
        'section_to_lines': s_to_l,
        'method_to_lines': m_to_l,
        'method_to_sections': m_to_s,
        'company_method_to_sections': cm_to_s,
        'company_method_to_lines': cm_to_l,
        'section_method_to_lines': sm_to_l,
        'company_section_method_to_lines': csm_to_l,
        'all_companies': sorted(all_companies),
        'all_sections': sorted(all_sections),
        'all_lines': sorted(all_lines),
        'all_methods': sorted(all_methods),
    }


def get_sections_for_company(mapping, company, method=None):
    """Return list of sections for a company and optional method."""
    has_company = bool(company and company not in ('전체', '(전체)', '전체 업체', '미지정'))
    has_method = bool(method and method not in ('전체', '(전체)', '전체 종류', '미지정'))
    m_key = method.upper() if has_method else None

    if has_company and has_method:
        secs = mapping.get('company_method_to_sections', {}).get((company, m_key), [])
        if secs:
            return sorted(secs)

    if has_company:
        secs = mapping.get('company_to_sections', {}).get(company, [])
    elif has_method:
        secs = mapping.get('method_to_sections', {}).get(m_key, mapping.get('all_sections', []))
    else:
        secs = mapping.get('all_sections', [])

    if has_method:
        valid_secs = set(mapping.get('method_to_sections', {}).get(m_key, []))
        if valid_secs:
            secs = [s for s in secs if s in valid_secs]

    return sorted(secs)


def get_lines_for_company_and_section(mapping, company, section, method=None):
    """Return list of lines for company, section and optional method."""
    has_company = bool(company and company not in ('전체', '(전체)', '전체 업체', '미지정'))
    has_section = bool(section and section not in ('전체', '(전체)', '전체 구간', '미지정'))
    has_method = bool(method and method not in ('전체', '(전체)', '전체 종류', '미지정'))
    m_key = method.upper() if has_method else None

    if has_company and has_section and has_method:
        lines = mapping.get('company_section_method_to_lines', {}).get((company, section, m_key), [])
        if lines:
            return sorted(lines)
        lines = mapping.get('section_method_to_lines', {}).get((section, m_key), [])
        if lines:
            return sorted(lines)

    if has_company and has_section:
        lines = mapping.get('company_section_to_lines', {}).get((company, section), [])
        if not lines:
            lines = mapping.get('section_to_lines', {}).get(section, [])
    elif has_section:
        if has_method:
            lines = mapping.get('section_method_to_lines', {}).get((section, m_key), [])
            if not lines:
                lines = mapping.get('section_to_lines', {}).get(section, [])
        else:
            lines = mapping.get('section_to_lines', {}).get(section, [])
    elif has_company:
        if has_method:
            lines = mapping.get('company_method_to_lines', {}).get((company, m_key), [])
            if not lines:
                co_sections = mapping.get('company_to_sections', {}).get(company, [])
                lines = set()
                for s in co_sections:
                    lines.update(mapping.get('company_section_to_lines', {}).get((company, s), []))
                lines.update(mapping.get('company_section_to_lines', {}).get((company, ''), []))
        else:
            co_sections = mapping.get('company_to_sections', {}).get(company, [])
            lines = set()
            for s in co_sections:
                lines.update(mapping.get('company_section_to_lines', {}).get((company, s), []))
            lines.update(mapping.get('company_section_to_lines', {}).get((company, ''), []))
    else:
        if has_method:
            lines = mapping.get('method_to_lines', {}).get(m_key, mapping.get('all_lines', []))
        else:
            lines = mapping.get('all_lines', [])

    if has_method:
        valid_lines = set(mapping.get('method_to_lines', {}).get(m_key, []))
        if valid_lines:
            lines = [l for l in lines if l in valid_lines]

    return sorted(list(lines))

