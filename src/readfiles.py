"""Read the supported resume tables without modifying the source package."""
from dataclasses import dataclass, field
from pathlib import Path
import re
from zipfile import ZipFile, BadZipFile
from xml.etree import ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


class ResumeReadError(ValueError):
    """The input cannot be interpreted safely as project records."""


@dataclass
class Project:
    name: str = ""
    period: str = ""
    overview: str = ""
    duties: str = ""
    technologies: str = ""
    source: str = ""
    stages: str = ""


@dataclass
class Extraction:
    projects: list[Project] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    names: list[str] = field(default_factory=list)


def cell_text(cell):
    paragraphs = []
    for paragraph in cell.findall("w:p", NS):
        parts = []
        for node in paragraph.iter():
            if node.tag == W + "t":
                parts.append(node.text or "")
            elif node.tag in (W + "br", W + "cr"):
                parts.append("\n")
            elif node.tag == W + "tab":
                parts.append("\t")
        paragraphs.append("".join(parts))
    return "\n".join(paragraphs).strip()


def compact(text):
    return re.sub(r"\s+", "", text)


def split_content(text):
    sections = {"name": [], "overview": [], "duties": []}
    current = "name"
    for line in text.splitlines():
        match = re.match(r"^\s*[【\[]\s*(概要|担当業務)\s*[】\]]\s*(.*)$", line)
        if match:
            current = "overview" if match[1] == "概要" else "duties"
            sections[current].append(match[2])
        else:
            sections[current].append(line)
    return ["\n".join(sections[key]).strip() for key in ("name", "overview", "duties")]


def read_projects(path: str | Path) -> Extraction:
    path = Path(path)
    if path.suffix.lower() != ".docx":
        raise ResumeReadError("入力には .docx ファイルを指定してください。")
    try:
        with ZipFile(path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
    except (BadZipFile, KeyError, ET.ParseError) as exc:
        raise ResumeReadError("Wordファイルを読み取れません。破損・暗号化・形式を確認してください。") from exc
    result = Extraction()
    # Read labeled names anywhere in the body; never use author metadata as a name.
    for p in root.findall(".//w:p", NS):
        text = "".join(p.itertext())
        match = re.match(r"^\s*(?:氏名|名前|ふりがな|フリガナ)\s*[:：]\s*(.+?)\s*$", text)
        if match:
            result.names.append(match[1])
    for table in root.findall(".//w:tbl", NS):
        rows = table.findall("w:tr", NS)
        for index, row in enumerate(rows):
            cells = row.findall("w:tc", NS)
            labels = [compact(cell_text(c)) for c in cells]
            if len(labels) == 2 and labels[0] in ("氏名", "名前", "ふりがな", "フリガナ"):
                result.names.append(cell_text(cells[1]))
    tables = root.findall("w:body/w:tbl", NS)
    all_tables = root.findall(".//w:tbl", NS)
    if len(all_tables) > len(tables):
        result.warnings.append(f"本文直下以外の表が{len(all_tables) - len(tables)}件あります。抽出対象外のため原文との照合が必要です。")
    seen = set()
    body_children = list(root.find("w:body", NS))
    previous_table = None
    previous_columns = None
    previous_width = None
    for number, table in enumerate(tables, 1):
        columns = None
        skipped_before_header = 0
        position = body_children.index(table)
        adjacent = position > 0 and body_children[position - 1] is previous_table
        rows = table.findall("w:tr", NS)
        # A continuation must be adjacent, rectangular and explicitly contain
        # dated project records. Never carry a schema across company paragraphs.
        if adjacent and previous_columns and rows:
            candidates = [row.findall("w:tc", NS) for row in rows]
            rectangular = all(len(cells) == previous_width for cells in candidates)
            if rectangular:
                dated = all(re.match(r"^\s*[0-9０-９]{4}\s*年\s*[0-9０-９]{1,2}\s*月\s*[～〜~－–—-]", cell_text(cells[previous_columns["period"]])) for cells in candidates)
                marked = all(re.search(r"[【\[]\s*(概要|担当業務)\s*[】\]]", cell_text(cells[previous_columns["content"]])) for cells in candidates)
                if dated and marked:
                    columns = previous_columns.copy()
                    width = previous_width
                    result.warnings.append(f"表{number}: 隣接する前表の列見出しを引き継ぎ、{len(rows)}行を継続案件として抽出しました。原文と照合してください。")
        for row_number, row in enumerate(table.findall("w:tr", NS), 1):
            location = f"表{number} 行{row_number}"
            cells = row.findall("w:tc", NS)
            texts = [cell_text(c) for c in cells]
            labels = [compact(t) for t in texts]
            if "期間" in labels and any("プロジェクト" in t or t in ("案件名", "案件名・内容") for t in labels):
                columns = {}
                for i, label in enumerate(labels):
                    if label == "期間": columns["period"] = i
                    elif "プロジェクト" in label or label in ("案件名", "案件名・内容"): columns["content"] = i
                    elif label == "担当業務": columns["duties"] = i
                    elif "環境" in label or "技術" in label: columns["technologies"] = i
                width = len(cells)
                continue
            if not any(texts):
                continue
            if columns is None:
                skipped_before_header += 1
                continue
            merged = any(c.find("w:tcPr/w:vMerge", NS) is not None or c.find("w:tcPr/w:gridSpan", NS) is not None for c in cells)
            if merged or len(cells) != width or row.find(".//w:tbl", NS) is not None:
                raise ResumeReadError(f"{location}: 結合セル・入れ子・列数不一致のため案件境界を判断できません。表を1行1案件に分割してください。")
            def value(key):
                return texts[columns[key]] if key in columns else ""
            name, overview, duties = split_content(value("content"))
            extra_duties = value("duties")
            project = Project(name, value("period"), overview, duties,
                              value("technologies"), location, extra_duties)
            if not name:
                result.warnings.append(f"{location}: 案件名が空欄です。前行の情報は補完しません。")
            key = (project.name, project.period, project.overview, project.duties, project.technologies, project.stages)
            if key in seen:
                result.warnings.append(f"{location}: 同一内容の案件があります。削除せず保持しました。")
            seen.add(key)
            lines = [compact(t) for t in (project.duties + "\n" + project.stages).splitlines() if compact(t)]
            if len(lines) != len(set(lines)):
                result.warnings.append(f"{location}: 担当業務に重複があります。削除せず保持しました。")
            result.projects.append(project)
        if columns is None:
            result.warnings.append(f"表{number}: 対応する列見出しがないため対象外です（空でない行: {skipped_before_header}行）。")
        elif skipped_before_header:
            result.warnings.append(f"表{number}: 最初の見出しより前の{skipped_before_header}行は抽出対象外です。原文を確認してください。")
        previous_table = table
        previous_columns = columns.copy() if columns is not None else None
        previous_width = width if columns is not None else None
    if not result.projects:
        raise ResumeReadError("対応するプロジェクト表が見つかりません。列見出しと1行1案件の構造を確認してください。")
    return result
