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
    all_companies = set()
    all_sections = set()
    all_lines = set()

    for date_str, day_data in (history_data or {}).items():
        if not isinstance(day_data, dict):
            continue
        for row in day_data.get('ndt_results', []):
            if not isinstance(row, dict):
                continue
            company = str(row.get('업체', '') or '').strip()
            section = str(row.get('구간', '') or '').strip()
            line = str(row.get('라인번호', '') or '').strip()

            if company:
                all_companies.add(company)
            if section:
                all_sections.add(section)
            if line:
                all_lines.add(line)

            if company and section:
                company_to_sections.setdefault(company, set()).add(section)
            if section and line:
                section_to_lines.setdefault(section, set()).add(line)
            if company and section and line:
                company_section_to_lines.setdefault((company, section), set()).add(line)
            elif company and line:
                company_section_to_lines.setdefault((company, ''), set()).add(line)

    c_to_s = {c: sorted(s) for c, s in company_to_sections.items()}
    cs_to_l = {k: sorted(v) for k, v in company_section_to_lines.items()}
    s_to_l = {s: sorted(lines) for s, lines in section_to_lines.items()}

    return {
        'company_to_sections': c_to_s,
        'company_section_to_lines': cs_to_l,
        'section_to_lines': s_to_l,
        'all_companies': sorted(all_companies),
        'all_sections': sorted(all_sections),
        'all_lines': sorted(all_lines),
    }


def get_sections_for_company(mapping, company):
    """Return list of sections for a company, or all sections if company is empty or '전체'."""
    if not company or company in ('전체', '(전체)', '전체 업체', '미지정'):
        return mapping.get('all_sections', [])
    return mapping.get('company_to_sections', {}).get(company, [])


def get_lines_for_company_and_section(mapping, company, section):
    """Return list of lines for company and section."""
    has_company = bool(company and company not in ('전체', '(전체)', '전체 업체', '미지정'))
    has_section = bool(section and section not in ('전체', '(전체)', '전체 구간', '미지정'))

    if has_company and has_section:
        lines = mapping.get('company_section_to_lines', {}).get((company, section), [])
        if lines:
            return lines
        return mapping.get('section_to_lines', {}).get(section, [])
    elif has_section:
        return mapping.get('section_to_lines', {}).get(section, [])
    elif has_company:
        co_sections = mapping.get('company_to_sections', {}).get(company, [])
        lines = set()
        for s in co_sections:
            lines.update(mapping.get('company_section_to_lines', {}).get((company, s), []))
        lines.update(mapping.get('company_section_to_lines', {}).get((company, ''), []))
        return sorted(lines)
    else:
        return mapping.get('all_lines', [])

