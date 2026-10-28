"""Run with python -m src.transfer input.docx."""
import argparse
from dataclasses import replace
from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import re
import sys
import unicodedata

from .readfiles import ResumeReadError, read_projects
from .formatting import format_project

DEFAULT_CONFIG = Path(__file__).resolve().parents[1] / "mask-config.local.json"
CATEGORIES = {"names": "氏名", "emails": "メールアドレス", "phones": "電話番号",
              "addresses": "住所", "birth_dates": "生年月日"}


class MaskConfigError(ValueError):
    """Invalid private configuration; never echo its contents."""


def load_mask_config(path=None):
    config_path = Path(path) if path is not None else DEFAULT_CONFIG
    if path is None and not config_path.exists():
        return {}
    try:
        config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, ValueError) as exc:
        raise MaskConfigError("マスク設定を読み取れません。JSONの形式・存在・権限を確認してください。") from exc
    if not isinstance(config, dict) or set(config) - set(CATEGORIES):
        raise MaskConfigError("マスク設定には names / emails / phones / addresses / birth_dates のみ指定できます。")
    for values in config.values():
        if not isinstance(values, list) or any(not isinstance(v, str) or not v.strip() for v in values):
            raise MaskConfigError("マスク設定の各項目は空でない文字列の配列にしてください。")
    return config


def mask_text(text: str, names=(), config=None, detections=None) -> str:
    """Shared masking pipeline; optional detections are for terminal review only."""
    config = config or {}

    def record(category, value):
        if detections is not None:
            detections.append((category, value))
        return "[" + category + "]"

    literals = []
    for key, category in CATEGORIES.items():
        values = [*config.get(key, []), *(names if key == "names" else [])]
        for value in set(values):
            if value.strip():
                literals.append((value, key, category))
    for value, key, category in sorted(literals, key=lambda item: len(item[0]), reverse=True):
        pattern = (r"[\s\u3000]*".join(re.escape(c) for c in value if not c.isspace())
                   if key == "names" else re.escape(value))
        text = re.sub(pattern, lambda m: record(category, m[0]), text)

    # Explicit personal labels only; no company/project name detection.
    for label, category in [("氏名|名前|ふりがな|フリガナ", "氏名"),
                            ("現住所|住所", "住所"), ("生年月日|誕生日", "生年月日")]:
        pattern = r"(?m)^([ \t]*(?:" + label + r")[ \t]*[:：][ \t]*)([^\n]+)"
        def labeled(match):
            value = match[2]
            if value.strip() == "[" + category + "]":
                return match[0]
            return match[1] + record(category, value)
        text = re.sub(pattern, labeled, text)
    text = re.sub(r"[\w.!#$%&'*+/=?^`{|}~\-]+[@＠][\w\-]+(?:[.．][\w\-]+)+",
                  lambda m: record("メールアドレス", m[0]), text)
    # Normalize only for matching; unmatched source characters remain unchanged.
    normalized = "".join(unicodedata.normalize("NFKC", c) if len(unicodedata.normalize("NFKC", c)) == 1 else c for c in text)
    matches = list(re.finditer(r"(?<!\d)(?:\+81[\s\-‐‑–ー()（）]*|0)(?:[\s\-‐‑–ー()（）]*\d){9,10}(?!\d)", normalized))
    for match in reversed(matches):
        text = text[:match.start()] + record("電話番号", text[match.start():match.end()]) + text[match.end():]
    return text


def build_output_path(path: Path) -> Path:
    timestamp = datetime.now(timezone(timedelta(hours=9))).strftime("%Y%m%d_%H%M%S_%f")
    return path.with_name(f"resume_converted_{timestamp}.txt")


def convert(path: str | Path, names=(), *, config=None, confirm=None):
    """Save only after reviewing the final text and explicitly approving it."""
    path = Path(path)
    result = read_projects(path)
    output = build_output_path(path)
    if output.exists():
        raise FileExistsError
    detections = []
    blocks = []
    for project in result.projects:
        # Mask the original fields before presentation changes newlines or labels.
        fields = {
            field: mask_text(getattr(project, field), [*result.names, *names], config, detections)
            for field in ("name", "period", "overview", "duties", "stages", "technologies")
        }
        blocks.append(format_project(replace(project, **fields)))
    text = "\n\n".join(blocks) + "\n"
    result.warnings.append("マスク漏れの可能性があります。企業名・顧客名・案件名は自動判定でマスクしません。最終出力を確認してください。")
    for warning in result.warnings:
        print(f"警告: {warning}", file=sys.stderr)
    print(f"抽出件数: {len(result.projects)}件（全職種が対象。警告の対象外行・表も確認してください）")
    print("マスク対象の確認（保存予定テキスト内の検出値・端末表示のみ）:")
    unique = list(dict.fromkeys(detections))
    if not unique:
        print("検出されたマスク対象はありません。マスク漏れがないか確認してください。")
    for category, original in unique:
        print(f"  {category}: {original!r} → [{category}]")
    print("\n保存予定の全文:\n" + text)
    try:
        answer = (input if confirm is None else confirm)("上記の内容で保存しますか？ yes と入力した場合のみ保存します [yes/NO]: ")
    except (EOFError, KeyboardInterrupt):
        answer = ""
    if answer.strip().lower() != "yes":
        print("保存を中止しました。出力ファイルは作成していません。")
        return None, result.warnings
    with output.open("x", encoding="utf-8", newline="\n") as stream:
        stream.write(text)
    return output, result.warnings


def main(argv=None):
    parser = argparse.ArgumentParser(description="Word職務経歴書をローカルでテキストに変換")
    parser.add_argument("input", type=Path)
    parser.add_argument("--mask-name", action="append", default=[], help="マスクする人名（複数指定可）")
    parser.add_argument("--config", type=Path, help="マスク設定JSON（省略時: mask-config.local.json）")
    args = parser.parse_args(argv)
    try:
        output, warnings = convert(args.input, args.mask_name, config=load_mask_config(args.config))
    except FileExistsError:
        print("エラー: 出力ファイルが既に存在します。既存ファイルを移動してから再実行してください。", file=sys.stderr)
        return 1
    except (ResumeReadError, MaskConfigError, OSError) as exc:
        detail = str(exc) if isinstance(exc, (ResumeReadError, MaskConfigError)) else "ファイルへのアクセスに失敗しました。存在・権限・空き容量を確認してください。"
        print(f"エラー: {detail}", file=sys.stderr)
        return 1
    if output is not None:
        print(f"変換完了: {output.name}（入力ファイルと同じフォルダ）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
