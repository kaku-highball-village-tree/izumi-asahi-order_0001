# -- coding: utf-8 --
###############################################################
#
# AsahiOrderAreaStoreMappingMaker_Cmd.py
#
# pip install openpyxl
#
###############################################################

import csv
import os
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


TARGET_WORKSHEET_NAME: str = "本州マグロ(週間)"
SUPPORTED_EXTENSIONS: set[str] = {".xlsx"}


def write_error_text(pszOutputFileFullPath: str, pszErrorMessage: str) -> None:
    """エラーメッセージをUTF-8のテキストファイルへ上書き保存します。"""
    pszDirectoryFullPath: str = os.path.dirname(pszOutputFileFullPath)
    if pszDirectoryFullPath != "":
        os.makedirs(pszDirectoryFullPath, exist_ok=True)
    with open(pszOutputFileFullPath, mode="w", encoding="utf-8", newline="") as objFile:
        objFile.write(pszErrorMessage.rstrip("\n") + "\n")


def get_error_file_full_path(pszInputFileFullPath: str) -> str:
    """入力ファイルと同じフォルダーに作る_error.txtのパスを返します。"""
    pszDirectoryFullPath: str = os.path.dirname(os.path.abspath(pszInputFileFullPath))
    pszBaseNameWithoutExtension: str = os.path.splitext(
        os.path.basename(pszInputFileFullPath)
    )[0]
    return os.path.join(pszDirectoryFullPath, pszBaseNameWithoutExtension + "_error.txt")


def report_processing_error(
    pszInputFileFullPath: str, pszProcessName: str, pszDetailMessage: str
) -> None:
    """標準エラーと入力ファイル用_error.txtへ同じエラーを出力します。"""
    pszErrorMessage: str = (
        "処理結果: エラー\n"
        + "入力ファイル: "
        + os.path.abspath(pszInputFileFullPath)
        + "\n発生した処理: "
        + pszProcessName
        + "\nエラー内容: "
        + pszDetailMessage
        + "\n"
    )
    print(pszErrorMessage, file=sys.stderr, end="")
    try:
        write_error_text(get_error_file_full_path(pszInputFileFullPath), pszErrorMessage)
    except OSError as objException:
        print(
            "Error: _error.txtの保存にも失敗しました。Detail = " + str(objException),
            file=sys.stderr,
        )


def remove_old_error_file(pszInputFileFullPath: str) -> None:
    """正常終了後、以前の処理で作られた_error.txtがあれば削除します。"""
    pszErrorFileFullPath: str = get_error_file_full_path(pszInputFileFullPath)
    if os.path.exists(pszErrorFileFullPath):
        os.remove(pszErrorFileFullPath)


def validate_input_path(pszInputFileFullPath: str) -> str:
    """入力パスと拡張子を検証し、絶対パスを返します。"""
    pszAbsolutePath: str = os.path.abspath(pszInputFileFullPath)
    if not os.path.exists(pszAbsolutePath):
        raise ValueError("入力ファイルが見つかりません。Path = " + pszAbsolutePath)
    if not os.path.isfile(pszAbsolutePath):
        raise ValueError("入力パスがファイルではありません。Path = " + pszAbsolutePath)
    pszExtension: str = os.path.splitext(pszAbsolutePath)[1].lower()
    if pszExtension not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            "入力ファイルの拡張子は.xlsxではありません。Path = " + pszAbsolutePath
        )
    return pszAbsolutePath


def normalize_cell_value(objValue: object) -> object:
    """空セルを空文字へ変換し、それ以外のセル値は変更せず返します。"""
    if objValue is None:
        return ""
    return objValue


def is_empty_cell_value(objValue: object) -> bool:
    """末尾空行・空列の判定対象となる空セルかを返します。"""
    return objValue is None or objValue == ""


def get_target_worksheet(objWorkbook: Workbook) -> Worksheet:
    """名前が完全一致する対象シートを返します。"""
    if TARGET_WORKSHEET_NAME not in objWorkbook.sheetnames:
        raise ValueError(
            "対象シートが見つかりません。Sheet = " + TARGET_WORKSHEET_NAME
        )
    return objWorkbook[TARGET_WORKSHEET_NAME]


def read_excel_rows(pszInputFileFullPath: str) -> tuple[list[list[object]], int]:
    """対象シートのセル値を読み、末尾の完全空行・空列を除いて返します。"""
    objWorkbook: Workbook = load_workbook(pszInputFileFullPath, data_only=True)
    try:
        objWorksheet: Worksheet = get_target_worksheet(objWorkbook)
        iLastDataRow: int = 0
        iLastDataColumn: int = 0
        for objRow in objWorksheet.iter_rows(
            min_row=1,
            max_row=objWorksheet.max_row,
            min_col=1,
            max_col=objWorksheet.max_column,
        ):
            for objCell in objRow:
                if not is_empty_cell_value(objCell.value):
                    iLastDataRow = max(iLastDataRow, objCell.row)
                    iLastDataColumn = max(iLastDataColumn, objCell.column)
        if iLastDataRow == 0 or iLastDataColumn == 0:
            return [], 0
        listRows: list[list[object]] = []
        for objRow in objWorksheet.iter_rows(
            min_row=1,
            max_row=iLastDataRow,
            min_col=1,
            max_col=iLastDataColumn,
        ):
            listRows.append(
                [normalize_cell_value(objCell.value) for objCell in objRow]
            )
        return listRows, iLastDataColumn
    finally:
        objWorkbook.close()


def get_output_file_path(pszInputFileFullPath: str) -> Path:
    """_本州マグロ(週間).tsvの出力パスを返します。"""
    objInputPath: Path = Path(pszInputFileFullPath)
    return objInputPath.with_name(
        objInputPath.stem + "_" + TARGET_WORKSHEET_NAME + ".tsv"
    )


def create_temporary_path(objOutputPath: Path) -> Path:
    """出力先と同じフォルダーに一意な一時TSVパスを作ります。"""
    iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
        prefix=objOutputPath.stem + "_", suffix=".tsv", dir=objOutputPath.parent
    )
    os.close(iFileDescriptor)
    return Path(pszTemporaryPath)


def save_tsv_rows(objOutputPath: Path, listRows: list[list[object]]) -> None:
    """セル値をUTF-8 BOMなし、タブ区切り、CRLFのTSVへ保存します。"""
    with objOutputPath.open(mode="w", encoding="utf-8", newline="") as objFile:
        objWriter = csv.writer(objFile, delimiter="\t", lineterminator="\r\n")
        objWriter.writerows(listRows)


def process_input_file(pszInputFileFullPath: str) -> None:
    """対象Excelシートのセル値から調査用TSVを作成します。"""
    pszValidatedPath: str = validate_input_path(pszInputFileFullPath)
    listRows, iColumnCount = read_excel_rows(pszValidatedPath)
    objOutputPath: Path = get_output_file_path(pszValidatedPath)
    objTemporaryPath: Path = create_temporary_path(objOutputPath)
    try:
        save_tsv_rows(objTemporaryPath, listRows)
        os.replace(objTemporaryPath, objOutputPath)
    finally:
        if objTemporaryPath.exists():
            objTemporaryPath.unlink()
    remove_old_error_file(pszValidatedPath)
    print("朝日注文エリア店舗対応調査用TSVファイルを作成しました。")
    print("Input: " + pszValidatedPath)
    print("Worksheet: " + TARGET_WORKSHEET_NAME)
    print("TSV: " + str(objOutputPath))
    print("Rows: " + str(len(listRows)))
    print("Columns: " + str(iColumnCount))


def main() -> int:
    """引数を確認して処理し、成功0・失敗1の終了コードを返します。"""
    if len(sys.argv) != 2:
        pszScriptFileName: str = os.path.basename(__file__)
        pszErrorMessage: str = (
            "Error: 入力Excelファイルパスを1件指定してください。\n"
            + "Usage: python "
            + pszScriptFileName
            + " <input_xlsx_file_path>\n"
        )
        print(pszErrorMessage, file=sys.stderr, end="")
        pszErrorFileFullPath: str = (
            os.path.splitext(pszScriptFileName)[0] + "_error_argument.txt"
        )
        try:
            write_error_text(pszErrorFileFullPath, pszErrorMessage)
        except OSError as objException:
            print(
                "Error: 引数エラーファイルを保存できません。Detail = "
                + str(objException),
                file=sys.stderr,
            )
        return 1
    pszInputFileFullPath: str = sys.argv[1]
    try:
        process_input_file(pszInputFileFullPath)
    except Exception as objException:
        report_processing_error(
            pszInputFileFullPath,
            "朝日注文エリア店舗対応調査TSV作成処理",
            str(objException),
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
