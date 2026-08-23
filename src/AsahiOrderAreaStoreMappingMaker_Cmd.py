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
import re
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


TARGET_WORKSHEET_NAMES: tuple[str, str] = ("本州マグロ(週間)", "割り")
SUPPORTED_EXTENSIONS: set[str] = {".xlsx"}
MAX_BACKUP_NUMBER: int = 9999


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
    try:
        pszWarningFileFullPath: str = get_warning_file_full_path(
            pszInputFileFullPath
        )
        if os.path.exists(pszWarningFileFullPath):
            os.remove(pszWarningFileFullPath)
    except OSError as objException:
        print(
            "Error: 古い_warning.txtを削除できませんでした。Detail = "
            + str(objException),
            file=sys.stderr,
        )


def remove_old_error_file(pszInputFileFullPath: str) -> None:
    """正常終了後、以前の処理で作られた_error.txtがあれば削除します。"""
    pszErrorFileFullPath: str = get_error_file_full_path(pszInputFileFullPath)
    if os.path.exists(pszErrorFileFullPath):
        os.remove(pszErrorFileFullPath)


def get_warning_file_full_path(pszInputFileFullPath: str) -> str:
    """入力ファイルと同じフォルダーに作る_warning.txtのパスを返します。"""
    pszDirectoryFullPath: str = os.path.dirname(os.path.abspath(pszInputFileFullPath))
    pszBaseNameWithoutExtension: str = os.path.splitext(
        os.path.basename(pszInputFileFullPath)
    )[0]
    return os.path.join(
        pszDirectoryFullPath, pszBaseNameWithoutExtension + "_warning.txt"
    )


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


def read_worksheet_rows(objWorksheet: Worksheet) -> tuple[list[list[object]], int]:
    """シートのセル値を読み、末尾の完全空行・空列を除いて返します。"""
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
        listRows.append([normalize_cell_value(objCell.value) for objCell in objRow])
    return listRows, iLastDataColumn


def read_excel_worksheets(
    pszInputFileFullPath: str,
) -> tuple[dict[str, tuple[list[list[object]], int]], list[str]]:
    """対象シートを1回のExcel読込で取得し、存在しないシート名も返します。"""
    objWorkbook: Workbook = load_workbook(pszInputFileFullPath, data_only=True)
    try:
        dictWorksheetResults: dict[str, tuple[list[list[object]], int]] = {}
        listMissingWorksheetNames: list[str] = []
        for pszWorksheetName in TARGET_WORKSHEET_NAMES:
            if pszWorksheetName not in objWorkbook.sheetnames:
                listMissingWorksheetNames.append(pszWorksheetName)
                continue
            dictWorksheetResults[pszWorksheetName] = read_worksheet_rows(
                objWorkbook[pszWorksheetName]
            )
        return dictWorksheetResults, listMissingWorksheetNames
    finally:
        objWorkbook.close()


def get_output_file_path(
    pszInputFileFullPath: str, pszWorksheetName: str
) -> Path:
    """入力ファイル名とシート名からTSVの出力パスを返します。"""
    objInputPath: Path = Path(pszInputFileFullPath)
    return objInputPath.with_name(objInputPath.stem + "_" + pszWorksheetName + ".tsv")


def create_temporary_path(objOutputPath: Path) -> Path:
    """出力先と同じフォルダーに同じ拡張子の一時パスを作ります。"""
    iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
        prefix=objOutputPath.stem + "_",
        suffix=objOutputPath.suffix,
        dir=objOutputPath.parent,
    )
    os.close(iFileDescriptor)
    return Path(pszTemporaryPath)


def save_tsv_rows(objOutputPath: Path, listRows: list[list[object]]) -> None:
    """セル値をUTF-8 BOMなし、タブ区切り、CRLFのTSVへ保存します。"""
    with objOutputPath.open(mode="w", encoding="utf-8", newline="") as objFile:
        objWriter = csv.writer(objFile, delimiter="\t", lineterminator="\r\n")
        objWriter.writerows(listRows)


def find_backup_numbers(objOutputPath: Path) -> list[int]:
    """通常名TSVに対応する既存バックアップの4桁番号を返します。"""
    objPattern: re.Pattern[str] = re.compile(
        r"^" + re.escape(objOutputPath.name) + r"\.bk([0-9]{4})\.tsv$"
    )
    listBackupNumbers: list[int] = []
    for objCandidatePath in objOutputPath.parent.iterdir():
        objMatch: re.Match[str] | None = objPattern.fullmatch(objCandidatePath.name)
        if objMatch is None or not objCandidatePath.is_file():
            continue
        iBackupNumber: int = int(objMatch.group(1))
        if 1 <= iBackupNumber <= MAX_BACKUP_NUMBER:
            listBackupNumbers.append(iBackupNumber)
    return listBackupNumbers


def get_next_backup_path(objOutputPath: Path) -> Path:
    """既存最大番号の次となる.bk%04d.tsvパスを返します。"""
    listBackupNumbers: list[int] = find_backup_numbers(objOutputPath)
    iBackupNumber: int = 1 if not listBackupNumbers else max(listBackupNumbers) + 1
    if iBackupNumber > MAX_BACKUP_NUMBER:
        raise ValueError(
            "バックアップ番号が最大値9999に到達しています。Path = "
            + str(objOutputPath)
        )
    return objOutputPath.with_name(
        objOutputPath.name + f".bk{iBackupNumber:04d}.tsv"
    )


def build_warning_text(
    pszInputFileFullPath: str,
    listMissingWorksheetNames: list[str],
    dictOutputPaths: dict[str, Path],
    dictBackupPaths: dict[str, Path],
) -> str:
    """未検出シート、作成TSV、旧TSVバックアップを含む警告文を返します。"""
    listLines: list[str] = [
        "処理結果: 警告",
        "入力ファイル: " + os.path.abspath(pszInputFileFullPath),
        "未検出シート: " + ", ".join(listMissingWorksheetNames),
        "警告内容: 対象シートが見つからないため、このシートのTSVは作成しませんでした。",
    ]
    for pszWorksheetName in TARGET_WORKSHEET_NAMES:
        if pszWorksheetName in listMissingWorksheetNames:
            continue
        listLines.append("作成したTSV: " + str(dictOutputPaths[pszWorksheetName]))
    for pszWorksheetName in listMissingWorksheetNames:
        if pszWorksheetName not in dictBackupPaths:
            continue
        listLines.append(
            "旧TSVの変更前パス: " + str(dictOutputPaths[pszWorksheetName])
        )
        listLines.append(
            "旧TSVのバックアップパス: " + str(dictBackupPaths[pszWorksheetName])
        )
    return "\n".join(listLines) + "\n"


def replace_output_files(
    dictTemporaryPaths: dict[Path, Path],
    dictMissingOutputBackups: dict[Path, Path],
    objWarningPath: Path,
    objTemporaryWarningPath: Path | None,
) -> None:
    """TSV・警告を置換し、欠落シートの旧TSVを連番名へ変更します。"""
    setManagedPaths: set[Path] = set(dictTemporaryPaths.keys()) | {objWarningPath}
    dictRollbackPaths: dict[Path, Path] = {}
    listRenamedMissingOutputs: list[tuple[Path, Path]] = []
    try:
        for objManagedPath in setManagedPaths:
            if not objManagedPath.exists():
                continue
            objRollbackPath: Path = create_temporary_path(objManagedPath)
            shutil.copy2(objManagedPath, objRollbackPath)
            dictRollbackPaths[objManagedPath] = objRollbackPath
        for objOutputPath, objBackupPath in dictMissingOutputBackups.items():
            if objBackupPath.exists():
                raise FileExistsError(
                    "バックアップ先がすでに存在します。Path = " + str(objBackupPath)
                )
            os.rename(objOutputPath, objBackupPath)
            listRenamedMissingOutputs.append((objOutputPath, objBackupPath))
        for objOutputPath, objTemporaryPath in dictTemporaryPaths.items():
            os.replace(objTemporaryPath, objOutputPath)
        if objTemporaryWarningPath is None:
            if objWarningPath.exists():
                objWarningPath.unlink()
        else:
            os.replace(objTemporaryWarningPath, objWarningPath)
    except Exception:
        for objManagedPath in setManagedPaths:
            objRollbackPath = dictRollbackPaths.get(objManagedPath)
            if objRollbackPath is not None and objRollbackPath.exists():
                os.replace(objRollbackPath, objManagedPath)
            elif objManagedPath.exists():
                objManagedPath.unlink()
        for objOutputPath, objBackupPath in reversed(listRenamedMissingOutputs):
            if objBackupPath.exists():
                os.rename(objBackupPath, objOutputPath)
        raise
    finally:
        for objRollbackPath in dictRollbackPaths.values():
            if objRollbackPath.exists():
                objRollbackPath.unlink()


def process_input_file(pszInputFileFullPath: str) -> None:
    """存在する対象Excelシートのセル値から調査用TSVを作成します。"""
    pszValidatedPath: str = validate_input_path(pszInputFileFullPath)
    dictWorksheetResults, listMissingWorksheetNames = read_excel_worksheets(
        pszValidatedPath
    )
    dictOutputPaths: dict[str, Path] = {
        pszWorksheetName: get_output_file_path(pszValidatedPath, pszWorksheetName)
        for pszWorksheetName in TARGET_WORKSHEET_NAMES
    }
    dictBackupPaths: dict[str, Path] = {}
    for pszWorksheetName in listMissingWorksheetNames:
        objMissingOutputPath: Path = dictOutputPaths[pszWorksheetName]
        if objMissingOutputPath.exists():
            dictBackupPaths[pszWorksheetName] = get_next_backup_path(
                objMissingOutputPath
            )
    dictTemporaryPaths: dict[Path, Path] = {}
    objWarningPath: Path = Path(get_warning_file_full_path(pszValidatedPath))
    objTemporaryWarningPath: Path | None = None
    try:
        for pszWorksheetName, (listRows, _) in dictWorksheetResults.items():
            objOutputPath: Path = dictOutputPaths[pszWorksheetName]
            objTemporaryPath: Path = create_temporary_path(objOutputPath)
            dictTemporaryPaths[objOutputPath] = objTemporaryPath
            save_tsv_rows(objTemporaryPath, listRows)
        if dictWorksheetResults and listMissingWorksheetNames:
            objTemporaryWarningPath = create_temporary_path(objWarningPath)
            write_error_text(
                str(objTemporaryWarningPath),
                build_warning_text(
                    pszValidatedPath,
                    listMissingWorksheetNames,
                    dictOutputPaths,
                    dictBackupPaths,
                ),
            )
        replace_output_files(
            dictTemporaryPaths,
            {
                dictOutputPaths[pszWorksheetName]: objBackupPath
                for pszWorksheetName, objBackupPath in dictBackupPaths.items()
            },
            objWarningPath,
            objTemporaryWarningPath,
        )
    finally:
        for objTemporaryPath in dictTemporaryPaths.values():
            if objTemporaryPath.exists():
                objTemporaryPath.unlink()
        if objTemporaryWarningPath is not None and objTemporaryWarningPath.exists():
            objTemporaryWarningPath.unlink()
    if not dictWorksheetResults:
        raise ValueError(
            "対象シートが見つかりません。対象シート = "
            + ", ".join(TARGET_WORKSHEET_NAMES)
        )
    remove_old_error_file(pszValidatedPath)
    print("朝日注文エリア店舗対応調査用TSVファイルを作成しました。")
    print("Input: " + pszValidatedPath)
    for pszWorksheetName in TARGET_WORKSHEET_NAMES:
        if pszWorksheetName not in dictWorksheetResults:
            continue
        listRows, iColumnCount = dictWorksheetResults[pszWorksheetName]
        print("Worksheet: " + pszWorksheetName)
        print("TSV: " + str(dictOutputPaths[pszWorksheetName]))
        print("Rows: " + str(len(listRows)))
        print("Columns: " + str(iColumnCount))
    for pszWorksheetName in listMissingWorksheetNames:
        print("Warning: 対象シートが見つかりません。Sheet = " + pszWorksheetName)
        if pszWorksheetName in dictBackupPaths:
            print("Backup: " + str(dictBackupPaths[pszWorksheetName]))
    if listMissingWorksheetNames:
        print("Warning File: " + str(objWarningPath))


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
