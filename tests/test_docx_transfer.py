from hashlib import sha256
import re
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import pytest

from src.readfiles import ResumeReadError, read_projects, W
from src.transfer import convert, main, mask_text, load_mask_config, MaskConfigError


def make_docx(path, records, merge=False):
    root = ET.Element(W + "document")
    body = ET.SubElement(root, W + "body")
    p = ET.SubElement(body, W + "p")
    ET.SubElement(ET.SubElement(p, W + "r"), W + "t").text = "氏名：山田 太郎"
    table = ET.SubElement(body, W + "tbl")
    for values in [["期間", "プロジェクト名・内容", "担当業務", "環境・言語等", "役割・規模"], *records]:
        row = ET.SubElement(table, W + "tr")
        for text in values:
            cell = ET.SubElement(row, W + "tc")
            for line in text.split("\n"):
                p = ET.SubElement(cell, W + "p")
                ET.SubElement(ET.SubElement(p, W + "r"), W + "t").text = line
    if merge:
        cell = table.findall(W + "tr")[1].find(W + "tc")
        ET.SubElement(ET.SubElement(cell, W + "tcPr"), W + "vMerge")
    with ZipFile(path, "w") as package:
        package.writestr("word/document.xml", ET.tostring(root))


RECORDS = [
    ["2024年9月\n～\n2024年11月\n(3ヶ月)\n※12月は自宅待機", "AML/CFTシステム開発\n【概要】\n機能改修\n\n【担当業務】\n単体テスト\n山田太郎へ連絡 sample@example.com 090-1234-5678", "単体テスト\n結合テスト", "C-shell\nPL/SQL\nPython\nLinux\nOracle", "メンバ\n5名"],
    ["2025年1月～2025年7月", "Concurアフターサポート\n【概要】運用保守・追加開発\n【担当業務】\n所属ローコード開発／運用。", "運用保守", "Java:8,11\nSQL Server", "メンバ"],
]


def test_multiple_projects_preserve_source_and_mask(tmp_path):
    source = tmp_path / "職務経歴書.docx"
    make_docx(source, RECORDS)
    before = sha256(source.read_bytes()).digest()
    extracted = read_projects(source)
    assert len(extracted.projects) == 2
    assert extracted.projects[0].period == RECORDS[0][0]
    assert extracted.projects[0].technologies == RECORDS[0][3]
    assert extracted.projects[1].duties == "所属ローコード開発／運用。"
    assert extracted.projects[1].stages == "運用保守"
    output, warnings = convert(source, confirm=lambda _: "yes")
    text = output.read_text(encoding="utf-8")
    assert re.fullmatch(r"職務経歴書_converted_\d{8}_\d{6}_\d{6}\.txt", output.name)
    assert "山田" not in text and "sample@example.com" not in text and "090-1234-5678" not in text
    assert "[氏名]" in text and "[メールアドレス]" in text and "[電話番号]" in text
    assert text.count("単体テスト") == 2
    assert any("重複" in warning for warning in warnings)
    first, second, _ = text.split("---")
    assert "Concur" not in first and "Python" not in second
    assert sha256(source.read_bytes()).digest() == before
    second_output, _ = convert(source, confirm=lambda _: "yes")
    assert second_output != output
    assert second_output.read_text(encoding="utf-8") == text
    assert output.read_text(encoding="utf-8") == text


def test_duplicates_and_missing_sections(tmp_path):
    source = tmp_path / "resume.docx"
    row = ["", "案件", "", "", ""]
    make_docx(source, [row, row])
    result = read_projects(source)
    assert len(result.projects) == 2
    assert result.projects[0].overview == ""
    assert result.projects[0].period == ""
    assert any("同一内容" in warning for warning in result.warnings)


def test_ambiguous_merge_does_not_export(tmp_path):
    source = tmp_path / "resume.docx"
    make_docx(source, RECORDS, merge=True)
    assert main([str(source)]) == 1
    assert not list(tmp_path.glob("resume_converted*.txt"))


def test_corrupt_missing_and_unsupported(tmp_path, capsys):
    source = tmp_path / "bad.docx"
    source.write_bytes(b"invalid")
    assert main([str(source)]) == 1
    assert "破損" in capsys.readouterr().err
    assert main([str(tmp_path / "missing.docx")]) == 1
    assert "アクセス" in capsys.readouterr().err
    with pytest.raises(ResumeReadError):
        read_projects(tmp_path / "resume.doc")
    make_docx(source, [])
    with pytest.raises(ResumeReadError, match="見つかりません"):
        read_projects(source)


def test_mask_variants():
    text = "佐藤 花子 ＡＢＣ＠example．com ０９０－１２３４－５６７８ +81 (90) 1234 5678 Java:8,11 2024年9月"
    masked = mask_text(text, ["佐藤花子"])
    assert masked.count("[電話番号]") == 2
    assert "[メールアドレス]" in masked
    assert "[氏名]" in masked
    assert "Java:8,11 2024年9月" in masked


@pytest.mark.parametrize("answer", ["", "no", "y"])
def test_decline_never_saves(tmp_path, answer, capsys):
    source = tmp_path / "resume.docx"
    make_docx(source, RECORDS)
    before = source.read_bytes()
    output, _ = convert(source, confirm=lambda _: answer)
    assert output is None
    assert not list(tmp_path.glob("resume_converted*.txt"))
    assert source.read_bytes() == before
    terminal = capsys.readouterr().out
    assert "sample@example.com" in terminal
    assert "保存予定の全文" in terminal


def test_zero_detections_still_requires_approval(tmp_path, capsys):
    source = tmp_path / "resume.docx"
    make_docx(source, [["2024年", "株式会社サンプル 顧客株式会社 / 案件ABC", "", "Python", ""]])
    calls = []
    def reject(prompt):
        calls.append(prompt)
        assert not list(tmp_path.glob("resume_converted*.txt"))
        return "no"
    output, _ = convert(source, confirm=reject)
    assert output is None and len(calls) == 1
    assert "検出されたマスク対象はありません" in capsys.readouterr().out
    output, _ = convert(source, confirm=lambda _: "yes")
    assert "株式会社サンプル 顧客株式会社 / 案件ABC" in output.read_text(encoding="utf-8")


@pytest.mark.parametrize("error", [EOFError, KeyboardInterrupt])
def test_interrupted_approval_does_not_save(tmp_path, error):
    source = tmp_path / "resume.docx"
    make_docx(source, RECORDS)
    def interrupt(_):
        raise error
    output, _ = convert(source, confirm=interrupt)
    assert output is None
    assert not list(tmp_path.glob("resume_converted*.txt"))


def test_config_and_labeled_personal_information(tmp_path):
    import json
    config_path = tmp_path / "mask-config.local.json"
    config_path.write_text(json.dumps({"names": ["佐藤花子"], "addresses": ["東京都架空区1-2"],
                                       "birth_dates": ["1990/01/02"], "emails": ["特殊メール"],
                                       "phones": ["内線123"]}, ensure_ascii=False), encoding="utf-8")
    config = load_mask_config(config_path)
    text = "株式会社サンプル 案件ABC 佐藤 花子 東京都架空区1-2 1990/01/02 特殊メール 内線123\n住所：大阪府架空市3-4\n生年月日：平成2年1月2日"
    masked = mask_text(text, config=config)
    for original in ["佐藤", "東京都", "1990", "特殊メール", "内線123", "大阪府", "平成2年"]:
        assert original not in masked
    assert "株式会社サンプル 案件ABC" in masked
    assert masked.count("[住所]") == 2
    assert masked.count("[生年月日]") == 2
    config_path.write_text('{"names": "private name"}', encoding="utf-8")
    with pytest.raises(MaskConfigError):
        load_mask_config(config_path)


def test_cli_config_approval_and_errors(tmp_path, monkeypatch, capsys):
    source = tmp_path / "resume.docx"
    make_docx(source, [["", "案件\n【概要】東京都架空区1-2", "", "", ""]])
    config_path = tmp_path / "private.json"
    config_path.write_text('{"addresses": ["東京都架空区1-2"]}', encoding="utf-8")
    monkeypatch.setattr("builtins.input", lambda _: "yes")
    assert main([str(source), "--config", str(config_path)]) == 0
    output, = tmp_path.glob("resume_converted_*.txt")
    assert "[住所]" in output.read_text(encoding="utf-8")
    assert "東京都" not in output.read_text(encoding="utf-8")
    assert main([str(source), "--config", str(tmp_path / "missing.json")]) == 1
    assert "設定" in capsys.readouterr().err


@pytest.mark.parametrize("contents", ['[]', '{"unknown": []}', '{"names": [""]}', '{', '{"addresses": [5]}'])
def test_invalid_settings_never_save(tmp_path, contents, monkeypatch):
    source = tmp_path / "resume.docx"
    make_docx(source, RECORDS)
    config_path = tmp_path / "private.json"
    config_path.write_text(contents, encoding="utf-8")
    def unexpected_prompt(_):
        pytest.fail("Invalid configuration must fail before prompting")
    monkeypatch.setattr("builtins.input", unexpected_prompt)
    assert main([str(source), "--config", str(config_path)]) == 1
    assert not list(tmp_path.glob("resume_converted*.txt"))


def test_default_config_optional_and_loaded(tmp_path, monkeypatch):
    config_path = tmp_path / "mask-config.local.json"
    monkeypatch.setattr("src.transfer.DEFAULT_CONFIG", config_path)
    assert load_mask_config() == {}
    config_path.write_text('{"names": ["架空太郎"]}', encoding="utf-8-sig")
    assert load_mask_config() == {"names": ["架空太郎"]}


def test_race_during_confirmation_never_overwrites(tmp_path, monkeypatch):
    source = tmp_path / "resume.docx"
    make_docx(source, RECORDS)
    destination = source.with_name("resume_converted_20261028_120000_123456.txt")
    monkeypatch.setattr("src.transfer.build_output_path", lambda _: destination)
    def approve(_):
        destination.write_text("先に作成されたファイル", encoding="utf-8")
        return "yes"
    with pytest.raises(FileExistsError):
        convert(source, confirm=approve)
    assert destination.read_text(encoding="utf-8") == "先に作成されたファイル"


def test_timestamp_filename_uses_japan_time(tmp_path, monkeypatch):
    from datetime import datetime, timedelta, timezone
    from src.transfer import build_output_path
    class FixedDateTime:
        @staticmethod
        def now(tz):
            assert tz.utcoffset(None) == timedelta(hours=9)
            return datetime(2026, 10, 28, 12, 34, 56, 123456, tzinfo=tz)
    monkeypatch.setattr("src.transfer.datetime", FixedDateTime)
    assert build_output_path(tmp_path / "resume.docx").name == "resume_converted_20261028_123456_123456.txt"


def test_existing_timestamp_file_never_overwritten(tmp_path, monkeypatch):
    source = tmp_path / "resume.docx"
    make_docx(source, RECORDS)
    destination = tmp_path / "resume_converted_20261028_120000_123456.txt"
    destination.write_text("既存ファイル", encoding="utf-8")
    monkeypatch.setattr("src.transfer.build_output_path", lambda _: destination)
    with pytest.raises(FileExistsError):
        convert(source, confirm=lambda _: pytest.fail("Must not prompt on collision"))
    assert destination.read_text(encoding="utf-8") == "既存ファイル"
