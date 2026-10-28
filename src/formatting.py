"""Presentation only: do not infer dates, duties, stages or technologies."""
import re


def one_line(value):
    return " ".join(line.strip() for line in value.splitlines() if line.strip())


def format_period(value):
    # Match only explicit Japanese year/month endpoints. Never calculate months.
    month = r"([0-9０-９]{4})\s*年\s*([0-9０-９]{1,2})\s*月"
    pattern = r"^\s*" + month + r"\s*[～〜~－–—-]\s*" + month
    match = re.match(pattern, value)
    if not match:
        return one_line(value)
    start_year, start_month, end_year, end_month = map(int, match.groups())
    if not (1 <= start_month <= 12 and 1 <= end_month <= 12):
        return one_line(value)
    result = f"{start_year}年{start_month}月～{end_year}年{end_month}月"
    remainder = value[match.end():]
    duration = re.match(r"\s*[（(]\s*([0-9０-９]+)\s*[ヶかカケ]月\s*[）)]", remainder)
    if duration:
        result += f"（{int(duration[1])}ヶ月）"
        remainder = remainder[duration.end():]
    note = one_line(remainder)
    return result + (" " + note if note else "")


def format_technologies(value):
    # Preserve order, duplicates and punctuation within each original line.
    return "、".join(line.strip() for line in value.splitlines() if line.strip())


def format_project(project):
    return (f"【案件／プロジェクト名】\n{project.name}\n\n"
            f"期間：{format_period(project.period)}\n\n"
            f"【概要】\n{project.overview}\n\n"
            f"【担当業務】\n{project.duties}\n\n"
            f"【担当工程】\n{project.stages}\n\n"
            f"【技術要素】\n{format_technologies(project.technologies)}\n\n---")
