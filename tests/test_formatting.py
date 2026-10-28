from zipfile import ZipFile
from xml.etree import ElementTree as ET

import pytest

from src.formatting import format_period, format_project, format_technologies
from src.readfiles import NS, W, read_projects
from test_docx_transfer import make_docx, RECORDS


@pytest.mark.parametrize(("original", "expected"), [
    ("2024年9月\n～\n2024年11月\n（3ヶ月）", "2024年9月～2024年11月（3ヶ月）"),
    ("2024 年 09 月 ～ 2025 年 01 月 ( 5 か月 )", "2024年9月～2025年1月（5ヶ月）"),
    ("２０２４年９月〜２０２４年１１月（３ヶ月）", "2024年9月～2024年11月（3ヶ月）"),
    ("2024年9月～2024年11月", "2024年9月～2024年11月"),
    ("2024年9月～2024年11月（99ヶ月）", "2024年9月～2024年11月（99ヶ月）"),
    ("2025年11月\n～\n現在", "2025年11月 ～ 現在"),
    ("2024年9月\n～\n2024年11月\n(3ヶ月)\n※12月は自宅待機", "2024年9月～2024年11月（3ヶ月） ※12月は自宅待機"),
    ("", ""),
])
def test_period_formats_only_recorded_values(original, expected):
    assert format_period(original) == expected
    assert "\n" not in format_period(original)


def test_technologies_order_duplicates_and_internal_punctuation():
    assert format_technologies("Java:8,11\nSQL Server\n\nJava:8,11\n C-shell ") == "Java:8,11、SQL Server、Java:8,11、C-shell"
    assert format_technologies("") == ""


def test_explanation_and_stages_preserve_repeated_words(tmp_path):
    source=tmp_path/"resume.docx"
    make_docx(source, RECORDS)
    projects=read_projects(source).projects
    p=projects[0]
    assert "単体テスト" in p.duties and "単体テスト" in p.stages
    text=format_project(p)
    assert text.count("単体テスト") == 2
    assert "【担当工程】\n単体テスト、結合テスト\n\n" in text
    assert "期間：2024年9月～2024年11月（3ヶ月） ※12月は自宅待機" in text
    assert "【技術要素】\nC-shell、PL/SQL、Python、Linux、Oracle" in text
    # Before/after preserve project count and the exact duty strings in order.
    assert len(projects) == 2
    for p in projects:
        before_duties="\n".join(t for t in (p.duties,p.stages) if t)
        assert before_duties == RECORDS[projects.index(p)][1].split("【担当業務】\n")[1] + "\n" + RECORDS[projects.index(p)][2]


def split_continuation(source, separator=False):
    with ZipFile(source) as package:
        root=ET.fromstring(package.read("word/document.xml"))
    body=root.find("w:body",NS)
    first=body.find("w:tbl",NS)
    rows=first.findall("w:tr",NS)
    for row in rows[2:]:
        first.remove(row)
        if separator:
            paragraph=ET.SubElement(body,W+"p")
            ET.SubElement(ET.SubElement(paragraph,W+"r"),W+"t").text="別会社の職歴"
        table=ET.SubElement(body,W+"tbl")
        table.append(row)
    with ZipFile(source,"w") as package:
        package.writestr("word/document.xml",ET.tostring(root))


def test_adjacent_headerless_tables_and_non_it_jobs(tmp_path):
    source=tmp_path/"resume.docx"
    make_docx(source,[RECORDS[0],
        ["2025年11月\n～\n2026年1月\n（3ヶ月）","製品改修\n【概要】原文\n【担当業務】探索テスト\n【成果】原文の成果","改修\nテスト","Java\nSQL",""],
        ["2026年2月～2026年8月（7ヶ月）","総務人事\n【概要】採用運用\n【担当業務】応募者管理","その他","Excel", ""],
        ["2000年3月～2002年6月（28ヶ月）","経理事務\n【概要】事務\n【担当業務】原文","","", ""],
        ["2003年3月～2015年12月（154ヶ月）","農業\n【概要】農作業\n【担当業務】原文","","", ""],
    ])
    split_continuation(source)
    result=read_projects(source)
    assert len(result.projects)==5
    assert [p.source for p in result.projects] == ["表1 行2","表2 行1","表3 行1","表4 行1","表5 行1"]
    assert [p.name for p in result.projects][1:]==["製品改修","総務人事","経理事務","農業"]
    assert "【成果】原文の成果" in result.projects[1].duties
    assert len([w for w in result.warnings if "引き継ぎ" in w])==4


def test_never_inherits_columns_across_company_paragraph(tmp_path):
    source=tmp_path/"resume.docx"
    make_docx(source,RECORDS)
    split_continuation(source,separator=True)
    result=read_projects(source)
    assert len(result.projects)==1
    assert any("表2" in w and "対象外" in w for w in result.warnings)


def test_headerless_unknown_shape_is_not_inferred(tmp_path):
    source=tmp_path/"resume.docx"
    make_docx(source,[RECORDS[0],["2025年11月～2026年1月", "内容見出しのない文章", "", "", ""]])
    split_continuation(source)
    result=read_projects(source)
    assert len(result.projects)==1
    assert any("表2" in w and "対象外" in w for w in result.warnings)


@pytest.mark.parametrize(("stages", "expected"), [
    ("基本設計\n製造\n単体テスト\n結合テスト", "基本設計、製造、単体テスト、結合テスト"),
    ("単体テスト\n\n 結合テスト \n単体テスト", "単体テスト、結合テスト、単体テスト"),
    ("運用保守", "運用保守"),
    ("", ""),
])
def test_stages_on_one_line(stages, expected):
    from src.readfiles import Project
    text = format_project(Project(stages=stages))
    actual = text.split("【担当工程】\n", 1)[1].split("\n\n【技術要素】", 1)[0]
    assert actual == expected
