"""JSONファイルを日本語が読める形式に整形するコマンドラインツール。"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Sequence


USAGE = "使用方法：py json_readable.py 入力ファイル.json"


class JsonReadableError(Exception):
    """利用者へ表示するエラーを表す。"""


def get_input_path(arguments: Sequence[str]) -> Path:
    """コマンドライン引数を確認し、入力ファイルのパスを返す。"""
    if not arguments:
        raise JsonReadableError(
            f"エラー：入力JSONファイルを指定してください。\n{USAGE}"
        )
    if len(arguments) > 1:
        raise JsonReadableError(
            f"エラー：指定できる入力ファイルは1つだけです。\n{USAGE}"
        )

    return Path(arguments[0])


def validate_input_file(input_path: Path) -> None:
    """入力パスの存在、ファイル種別、拡張子を確認する。"""
    if not input_path.exists():
        raise JsonReadableError(
            "エラー：入力ファイルが見つかりません。\n"
            f"入力ファイル：{input_path}"
        )
    if not input_path.is_file():
        raise JsonReadableError("エラー：指定されたパスはファイルではありません。")
    if input_path.suffix.lower() != ".json":
        raise JsonReadableError("エラー：JSONファイルを指定してください。")


def create_output_path(input_path: Path) -> Path:
    """入力ファイルと同じフォルダーに出力パスを作る。"""
    return input_path.with_name(f"{input_path.stem}_readable.json")


def validate_output_file(output_path: Path) -> None:
    """既存の出力ファイルを上書きしないよう確認する。"""
    if output_path.exists():
        raise JsonReadableError(
            "エラー：出力ファイルはすでに存在します。\n"
            f"出力先：{output_path}"
        )


def load_json_file(input_path: Path) -> Any:
    """UTF-8（BOMの有無を問わない）でJSONファイルを読み込む。"""
    try:
        with input_path.open("r", encoding="utf-8-sig") as input_file:
            return json.load(input_file)
    except UnicodeDecodeError as error:
        raise JsonReadableError(
            "エラー：入力ファイルをUTF-8として読み込めません。"
        ) from error
    except json.JSONDecodeError as error:
        raise JsonReadableError(
            "エラー：JSONの形式が正しくありません。\n"
            f"行：{error.lineno}\n"
            f"列：{error.colno}\n"
            f"詳細：{error.msg}"
        ) from error
    except OSError as error:
        raise JsonReadableError(
            "エラー：ファイルの読み書き中に問題が発生しました。\n"
            f"詳細：{error}"
        ) from error


def save_readable_json(data: Any, output_path: Path) -> None:
    """JSONを日本語のまま2文字で字下げし、新しいファイルへ保存する。"""
    created = False
    try:
        # 排他的作成モードを使い、確認後に作られた同名ファイルも上書きしない。
        with output_path.open("x", encoding="utf-8", newline="\n") as output_file:
            created = True
            json.dump(data, output_file, ensure_ascii=False, indent=2)
            output_file.write("\n")
    except FileExistsError as error:
        raise JsonReadableError(
            "エラー：出力ファイルはすでに存在します。\n"
            f"出力先：{output_path}"
        ) from error
    except (OSError, UnicodeEncodeError) as error:
        # 書き込み途中の不完全なファイルは残さない。
        if created:
            try:
                output_path.unlink()
            except OSError:
                pass
        raise JsonReadableError(
            "エラー：出力ファイルを作成できません。\n"
            f"出力先：{output_path}\n"
            f"詳細：{error}"
        ) from error


def display_json_summary(data: Any) -> None:
    """最上位のJSON型と、該当する場合は要素数を表示する。"""
    if isinstance(data, dict):
        print("JSONの種類：オブジェクト")
        print(f"最上位の項目数：{len(data)}")

        items = data.get("items")
        if isinstance(items, list):
            print(f"itemsの要素数：{len(items)}")
            if "count" in data and data["count"] != len(items):
                print("警告：countの値とitemsの要素数が一致しません。")
    elif isinstance(data, list):
        print("JSONの種類：配列")
        print(f"最上位の要素数：{len(data)}")
    elif isinstance(data, str):
        print("JSONの種類：文字列")
    elif isinstance(data, bool):
        print("JSONの種類：真偽値")
    elif data is None:
        print("JSONの種類：null")
    else:
        print("JSONの種類：数値")


def main() -> None:
    """引数の確認からJSONの保存まで、プログラム全体を実行する。"""
    try:
        input_path = get_input_path(sys.argv[1:])
        validate_input_file(input_path)

        output_path = create_output_path(input_path)
        validate_output_file(output_path)

        data = load_json_file(input_path)
        save_readable_json(data, output_path)

        print("JSONファイルを読み込みました。")
        print(f"入力ファイル：{input_path}")
        display_json_summary(data)
        print("出力ファイルを作成しました。")
        print(f"出力先：{output_path}")
    except JsonReadableError as error:
        print(error)
        sys.exit(1)


if __name__ == "__main__":
    main()
