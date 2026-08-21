# -- coding: utf-8 --
###############################################################
#
# AsahiOrderTemplateMaker_Cmd.py
#
# pip install openpyxl
#
###############################################################

import csv
import os
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


EXPECTED_HEADERS: tuple[str, str, str] = ("productCode", "productName", "spec")
SUPPORTED_EXTENSIONS: set[str] = {".xlsx", ".tsv", ".csv"}


class ProductRow:
    """テンプレートへ出力する1商品の3列を保持します。"""

    def __init__(self, product_code: str, product_name: str, spec: str) -> None:
        self.product_code: str = product_code
        self.product_name: str = product_name
        self.spec: str = spec


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
        raise ValueError("入力ファイルの拡張子は未対応です。Path = " + pszAbsolutePath)
    return pszAbsolutePath


def normalize_header(objValue: object) -> str:
    """ヘッダー比較用に値を前後空白のない文字列へ変換します。"""
    if objValue is None:
        return ""
    return str(objValue).strip()


def normalize_text(objValue: object) -> str:
    """出力用にNoneを空文字、それ以外を文字列へ変換します。"""
    if objValue is None:
        return ""
    return str(objValue)


def validate_headers(listValues: list[object], pszSourceName: str) -> None:
    """先頭3列が仕様どおりのヘッダーであることを確認します。"""
    if len(listValues) < 3:
        raise ValueError(pszSourceName + "のヘッダーは3列未満です。")
    tupleHeaders: tuple[str, str, str] = tuple(
        normalize_header(objValue) for objValue in listValues[:3]
    )  # type: ignore[assignment]
    if tupleHeaders != EXPECTED_HEADERS:
        raise ValueError(
            pszSourceName
            + "の先頭3列はproductCode、productName、specではありません。"
        )


def find_target_worksheet(objWorkbook: Workbook) -> Worksheet:
    """A1～C1が仕様どおりのシートを探し、1枚だけなら返します。"""
    listTargetWorksheets: list[Worksheet] = []
    for objWorksheet in objWorkbook.worksheets:
        listHeaders: list[object] = [
            objWorksheet.cell(row=1, column=iColumn).value for iColumn in range(1, 4)
        ]
        if tuple(normalize_header(objValue) for objValue in listHeaders) == EXPECTED_HEADERS:
            listTargetWorksheets.append(objWorksheet)
    if len(listTargetWorksheets) == 0:
        raise ValueError(
            "A1～C1がproductCode、productName、specのシートが見つかりません。"
        )
    if len(listTargetWorksheets) > 1:
        raise ValueError(
            "対象シートが複数見つかりました。対象シート = "
            + ", ".join(objWorksheet.title for objWorksheet in listTargetWorksheets)
        )
    return listTargetWorksheets[0]


def build_product_row(listValues: list[object], iRow: int) -> ProductRow | None:
    """先頭3列を検証し、空行ならNone、それ以外なら商品行を返します。"""
    if len(listValues) < 3:
        if all(normalize_text(objValue).strip() == "" for objValue in listValues):
            return None
        raise ValueError(str(iRow) + "行目は3列未満です。")
    listTexts: list[str] = [normalize_text(objValue) for objValue in listValues[:3]]
    if all(pszValue.strip() == "" for pszValue in listTexts):
        return None
    pszProductCode: str = listTexts[0].strip()
    if pszProductCode == "":
        raise ValueError(
            str(iRow) + "行目はproductCodeが空ですが、ほかの対象列に値があります。"
        )
    return ProductRow(pszProductCode, listTexts[1], listTexts[2])


def validate_unique_product_codes(
    listProductRows: list[ProductRow], pszSourceName: str
) -> None:
    """productCodeが重複していないことを確認します。"""
    setSeenCodes: set[str] = set()
    setDuplicateCodes: set[str] = set()
    for objProductRow in listProductRows:
        if objProductRow.product_code in setSeenCodes:
            setDuplicateCodes.add(objProductRow.product_code)
        setSeenCodes.add(objProductRow.product_code)
    if setDuplicateCodes:
        raise ValueError(
            pszSourceName
            + "でproductCodeが重複しています。productCode = "
            + ", ".join(sorted(setDuplicateCodes))
        )


def read_excel_rows(pszInputFileFullPath: str) -> list[ProductRow]:
    """Excelの対象シートからテンプレート用商品行を読み取ります。"""
    objWorkbook: Workbook = load_workbook(pszInputFileFullPath, data_only=True)
    objWorksheet: Worksheet = find_target_worksheet(objWorkbook)
    listProductRows: list[ProductRow] = []
    for iRow in range(2, objWorksheet.max_row + 1):
        listValues: list[object] = [
            objWorksheet.cell(row=iRow, column=iColumn).value for iColumn in range(1, 4)
        ]
        objProductRow: ProductRow | None = build_product_row(listValues, iRow)
        if objProductRow is not None:
            listProductRows.append(objProductRow)
    validate_unique_product_codes(listProductRows, "Excel")
    return listProductRows


def read_delimited_rows_with_encoding(
    pszInputFileFullPath: str, pszEncoding: str, pszDelimiter: str
) -> list[list[str]]:
    """指定文字コードと区切り文字で全レコードを読み取ります。"""
    with open(
        pszInputFileFullPath, mode="r", encoding=pszEncoding, newline=""
    ) as objFile:
        return list(csv.reader(objFile, delimiter=pszDelimiter, strict=True))


def read_delimited_rows(pszInputFileFullPath: str) -> list[ProductRow]:
    """CSVまたはTSVをUTF-8優先、失敗時CP932で読み取ります。"""
    pszExtension: str = os.path.splitext(pszInputFileFullPath)[1].lower()
    pszDelimiter: str = "\t" if pszExtension == ".tsv" else ","
    pszSourceName: str = "TSV" if pszExtension == ".tsv" else "CSV"
    try:
        listRows: list[list[str]] = read_delimited_rows_with_encoding(
            pszInputFileFullPath, "utf-8-sig", pszDelimiter
        )
    except UnicodeDecodeError:
        listRows = read_delimited_rows_with_encoding(
            pszInputFileFullPath, "cp932", pszDelimiter
        )
    if not listRows:
        raise ValueError(pszSourceName + "ファイルが空です。")
    validate_headers(listRows[0], pszSourceName)
    listProductRows: list[ProductRow] = []
    for iRow, listValues in enumerate(listRows[1:], start=2):
        objProductRow: ProductRow | None = build_product_row(listValues, iRow)
        if objProductRow is not None:
            listProductRows.append(objProductRow)
    validate_unique_product_codes(listProductRows, pszSourceName)
    return listProductRows


def get_output_file_paths(pszInputFileFullPath: str) -> tuple[Path, Path]:
    """_step0001.xlsxと_step0001.tsvの出力パスを返します。"""
    objInputPath: Path = Path(pszInputFileFullPath)
    objBasePath: Path = objInputPath.with_name(objInputPath.stem + "_step0001")
    return objBasePath.with_suffix(".xlsx"), objBasePath.with_suffix(".tsv")


def create_temporary_path(objOutputPath: Path, pszSuffix: str) -> Path:
    """出力先と同じフォルダーに一意な一時ファイルパスを作ります。"""
    iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
        prefix=objOutputPath.stem + "_", suffix=pszSuffix, dir=objOutputPath.parent
    )
    os.close(iFileDescriptor)
    return Path(pszTemporaryPath)


def save_excel_template(objOutputPath: Path, listProductRows: list[ProductRow]) -> None:
    """3列の新規Excelテンプレートを保存します。"""
    objWorkbook: Workbook = Workbook()
    objWorksheet: Worksheet = objWorkbook.active
    objWorksheet.append(list(EXPECTED_HEADERS))
    for objProductRow in listProductRows:
        objWorksheet.append(
            [objProductRow.product_code, objProductRow.product_name, objProductRow.spec]
        )
        objWorksheet.cell(row=objWorksheet.max_row, column=1).number_format = "@"
    objWorkbook.save(objOutputPath)


def save_tsv_template(objOutputPath: Path, listProductRows: list[ProductRow]) -> None:
    """3列のTSVをUTF-8 BOMなし、CRLFで保存します。"""
    with objOutputPath.open(mode="w", encoding="utf-8", newline="") as objFile:
        objWriter = csv.writer(objFile, delimiter="\t", lineterminator="\r\n")
        objWriter.writerow(EXPECTED_HEADERS)
        for objProductRow in listProductRows:
            objWriter.writerow(
                [objProductRow.product_code, objProductRow.product_name, objProductRow.spec]
            )


def replace_output_files(
    objTemporaryExcelPath: Path,
    objTemporaryTsvPath: Path,
    objExcelOutputPath: Path,
    objTsvOutputPath: Path,
) -> None:
    """2出力を置換し、失敗時は可能な限り以前の状態へ戻します。"""
    listOutputPaths: list[Path] = [objExcelOutputPath, objTsvOutputPath]
    listTemporaryPaths: list[Path] = [objTemporaryExcelPath, objTemporaryTsvPath]
    dictBackupPaths: dict[Path, Path] = {}
    listReplacedPaths: list[Path] = []
    try:
        for objOutputPath in listOutputPaths:
            if objOutputPath.exists():
                objBackupPath: Path = create_temporary_path(objOutputPath, ".backup")
                shutil.copy2(objOutputPath, objBackupPath)
                dictBackupPaths[objOutputPath] = objBackupPath
        for objTemporaryPath, objOutputPath in zip(listTemporaryPaths, listOutputPaths):
            os.replace(objTemporaryPath, objOutputPath)
            listReplacedPaths.append(objOutputPath)
    except Exception:
        for objOutputPath in reversed(listReplacedPaths):
            objBackupPath = dictBackupPaths.get(objOutputPath)
            if objBackupPath is not None and objBackupPath.exists():
                os.replace(objBackupPath, objOutputPath)
            elif objOutputPath.exists():
                objOutputPath.unlink()
        raise
    finally:
        for objBackupPath in dictBackupPaths.values():
            if objBackupPath.exists():
                objBackupPath.unlink()
        for objTemporaryPath in listTemporaryPaths:
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()


def process_input_file(pszInputFileFullPath: str) -> None:
    """入力を検証し、同じ内容のXLSXとTSVテンプレートを作成します。"""
    pszValidatedPath: str = validate_input_path(pszInputFileFullPath)
    pszExtension: str = os.path.splitext(pszValidatedPath)[1].lower()
    if pszExtension == ".xlsx":
        listProductRows: list[ProductRow] = read_excel_rows(pszValidatedPath)
    else:
        listProductRows = read_delimited_rows(pszValidatedPath)
    objExcelOutputPath, objTsvOutputPath = get_output_file_paths(pszValidatedPath)
    objTemporaryExcelPath: Path = create_temporary_path(objExcelOutputPath, ".xlsx")
    objTemporaryTsvPath: Path = create_temporary_path(objTsvOutputPath, ".tsv")
    try:
        save_excel_template(objTemporaryExcelPath, listProductRows)
        save_tsv_template(objTemporaryTsvPath, listProductRows)
        replace_output_files(
            objTemporaryExcelPath,
            objTemporaryTsvPath,
            objExcelOutputPath,
            objTsvOutputPath,
        )
    finally:
        for objTemporaryPath in (objTemporaryExcelPath, objTemporaryTsvPath):
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()
    remove_old_error_file(pszValidatedPath)
    print("朝日注文テンプレートファイルを作成しました。")
    print("Input: " + pszValidatedPath)
    print("Excel: " + str(objExcelOutputPath))
    print("TSV: " + str(objTsvOutputPath))
    print("Rows: " + str(len(listProductRows)))


def main() -> int:
    """引数を確認して処理し、成功0・失敗1の終了コードを返します。"""
    if len(sys.argv) != 2:
        pszScriptFileName: str = os.path.basename(__file__)
        pszErrorMessage: str = (
            "Error: 入力ファイルパスを1件指定してください。\n"
            + "Usage: python "
            + pszScriptFileName
            + " <input_file_path>\n"
        )
        print(pszErrorMessage, file=sys.stderr, end="")
        pszErrorFileFullPath: str = os.path.splitext(pszScriptFileName)[0] + "_error_argument.txt"
        try:
            write_error_text(pszErrorFileFullPath, pszErrorMessage)
        except OSError as objException:
            print(
                "Error: 引数エラーファイルを保存できません。Detail = " + str(objException),
                file=sys.stderr,
            )
        return 1
    pszInputFileFullPath: str = sys.argv[1]
    try:
        process_input_file(pszInputFileFullPath)
    except Exception as objException:
        report_processing_error(
            pszInputFileFullPath,
            "朝日注文テンプレートファイル作成処理",
            str(objException),
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
