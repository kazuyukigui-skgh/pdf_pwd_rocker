#!/usr/bin/env python3
"""
PDF Locker - コアロジック（共通処理）

GUIアプリ（Tkinter）とWebアプリ（Streamlit）の両方で使用される共通処理を提供します。
このファイルを修正すれば、両方のアプリに変更が反映されます。

主な機能:
- PDFへのパスワード設定（AES-256暗号化）
- Office文書（Word/Excel/PowerPoint）からPDFへの変換
- ファイル処理のユーティリティ
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
from typing import Tuple, Optional, List, BinaryIO
from dataclasses import dataclass

# pypdfのインポート
try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import PdfReadError
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False
    PdfReadError = Exception  # フォールバック

# Office文書変換用ライブラリ
# docx2pdf（Word用）
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

# comtypes（Excel/PowerPoint用・Windows専用）
if sys.platform == "win32":
    try:
        import comtypes.client
        COMTYPES_AVAILABLE = True
    except ImportError:
        COMTYPES_AVAILABLE = False
else:
    COMTYPES_AVAILABLE = False


# 対応ファイル拡張子
SUPPORTED_EXTENSIONS = {'.pdf', '.docx', '.xlsx', '.pptx'}
PDF_EXTENSION = '.pdf'
OFFICE_EXTENSIONS = {'.docx', '.xlsx', '.pptx'}


@dataclass
class ProcessResult:
    """処理結果を格納するデータクラス"""
    success: bool
    output_path: Optional[str] = None
    error_message: str = ""
    original_filename: str = ""


def check_dependencies() -> Tuple[bool, str]:
    """
    必要なライブラリがインストールされているかチェック

    Returns:
        (利用可能フラグ, エラーメッセージ)
    """
    if not PYPDF_AVAILABLE:
        return False, "pypdfライブラリが見つかりません。\npip install pypdf[crypto] を実行してください。"
    return True, ""


def is_supported_file(file_path: str) -> bool:
    """
    サポートされているファイル形式かどうかを判定

    Args:
        file_path: ファイルパス

    Returns:
        サポートされている場合True
    """
    ext = Path(file_path).suffix.lower()
    return ext in SUPPORTED_EXTENSIONS


def get_file_type_icon(file_path: str) -> str:
    """
    ファイル形式に応じたアイコンを取得

    Args:
        file_path: ファイルパス

    Returns:
        絵文字アイコン
    """
    ext = Path(file_path).suffix.lower()
    icon_map = {
        '.pdf': '📄',
        '.docx': '📝',
        '.xlsx': '📊',
        '.pptx': '📽️'
    }
    return icon_map.get(ext, '📁')


def convert_office_to_pdf(input_path: str, output_path: str) -> Tuple[bool, str]:
    """
    Office文書をPDFに変換する

    Args:
        input_path: 入力ファイルパス（.docx, .xlsx, .pptx）
        output_path: 出力PDFパス

    Returns:
        (成功フラグ, エラーメッセージ)
    """
    file_ext = Path(input_path).suffix.lower()

    # Word文書の変換
    if file_ext == '.docx':
        if DOCX2PDF_AVAILABLE:
            try:
                docx2pdf_convert(input_path, output_path)
                return True, ""
            except Exception as e:
                return False, f"Word文書の変換に失敗しました: {str(e)}"
        else:
            return False, "Word文書の変換機能が利用できません。\ndocx2pdfライブラリをインストールしてください。"

    # Excel/PowerPointの変換（Windows専用）
    elif file_ext in ['.xlsx', '.pptx']:
        if not sys.platform == "win32":
            return False, "Excel/PowerPoint変換はWindows専用です。"

        if not COMTYPES_AVAILABLE:
            return False, "Office変換機能が利用できません。\ncomtypesライブラリをインストールしてください。"

        try:
            if file_ext == '.xlsx':
                # Excel変換
                excel = comtypes.client.CreateObject('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(str(Path(input_path).absolute()))
                wb.ExportAsFixedFormat(0, str(Path(output_path).absolute()))
                wb.Close(False)
                excel.Quit()

                return True, ""

            elif file_ext == '.pptx':
                # PowerPoint変換
                powerpoint = comtypes.client.CreateObject('PowerPoint.Application')
                powerpoint.Visible = 1

                presentation = powerpoint.Presentations.Open(str(Path(input_path).absolute()))
                presentation.SaveAs(str(Path(output_path).absolute()), 32)  # 32 = ppSaveAsPDF
                presentation.Close()
                powerpoint.Quit()

                return True, ""

        except Exception as e:
            error_msg = str(e)
            if "Microsoft Office" in error_msg or "Excel" in error_msg or "PowerPoint" in error_msg:
                return False, f"{file_ext}の変換に失敗しました。\nMicrosoft Officeがインストールされているか確認してください。"
            return False, f"{file_ext}の変換に失敗しました: {error_msg}"

    return False, f"未対応の形式です: {file_ext}"


def validate_password(password: str) -> Tuple[bool, str]:
    """
    パスワードのバリデーション

    Args:
        password: チェックするパスワード

    Returns:
        (有効フラグ, エラーメッセージ)
    """
    if not password:
        return False, "パスワードを入力してください。"

    if len(password) < 4:
        return False, "パスワードは4文字以上にしてください。"

    return True, ""


def lock_pdf_bytes(pdf_bytes: bytes, password: str) -> Tuple[bool, bytes, str]:
    """
    PDFバイトデータにパスワードを設定

    Args:
        pdf_bytes: PDFのバイトデータ
        password: 設定するパスワード

    Returns:
        (成功フラグ, 暗号化されたPDFバイト, エラーメッセージ)
    """
    if not PYPDF_AVAILABLE:
        return False, b"", "pypdfライブラリが利用できません。"

    try:
        import io

        # PDFを読み込む
        reader = PdfReader(io.BytesIO(pdf_bytes))

        # 既に暗号化されている場合
        if reader.is_encrypted:
            return False, b"", "すでに鍵がかかっています"

        # 新しいPDFを作成
        writer = PdfWriter()

        # すべてのページをコピー
        for page in reader.pages:
            writer.add_page(page)

        # メタデータをコピー
        if reader.metadata:
            writer.add_metadata(reader.metadata)

        # AES-256で暗号化
        writer.encrypt(
            user_password=password,
            owner_password=password,
            algorithm="AES-256"
        )

        # バイトデータとして出力
        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        return True, output.getvalue(), ""

    except PdfReadError:
        return False, b"", "PDFファイルが壊れているかもしれません"
    except Exception as e:
        return False, b"", f"エラーが発生しました: {str(e)}"


def lock_pdf_file(input_path: str, output_path: str, password: str) -> Tuple[bool, str]:
    """
    PDFファイルにパスワードを設定してファイルに保存

    Args:
        input_path: 入力PDFファイルパス
        output_path: 出力PDFファイルパス
        password: 設定するパスワード

    Returns:
        (成功フラグ, エラーメッセージ)
    """
    if not PYPDF_AVAILABLE:
        return False, "pypdfライブラリが利用できません。"

    try:
        # PDFを読み込む
        reader = PdfReader(input_path)

        # 既に暗号化されている場合
        if reader.is_encrypted:
            return False, "すでに鍵がかかっています"

        # 新しいPDFを作成
        writer = PdfWriter()

        # すべてのページをコピー
        for page in reader.pages:
            writer.add_page(page)

        # メタデータをコピー
        if reader.metadata:
            writer.add_metadata(reader.metadata)

        # AES-256で暗号化
        writer.encrypt(
            user_password=password,
            owner_password=password,
            algorithm="AES-256"
        )

        # ファイルに保存
        with open(output_path, "wb") as f:
            writer.write(f)

        return True, ""

    except PdfReadError:
        return False, "PDFファイルが壊れているかもしれません"
    except PermissionError:
        return False, "このファイルは開けません（使用中の可能性）"
    except Exception as e:
        return False, f"エラーが発生しました: {str(e)}"


def process_file(
    file_path: str,
    password: str,
    output_dir: Optional[str] = None,
    output_prefix: str = "鍵付き_"
) -> ProcessResult:
    """
    ファイルを処理してパスワード付きPDFを作成

    Office文書の場合は自動的にPDFに変換してからパスワードを設定します。

    Args:
        file_path: 入力ファイルパス
        password: 設定するパスワード
        output_dir: 出力ディレクトリ（Noneの場合は入力ファイルと同じ場所）
        output_prefix: 出力ファイル名のプレフィックス

    Returns:
        ProcessResult: 処理結果
    """
    original_path = Path(file_path)
    file_ext = original_path.suffix.lower()

    # 出力ディレクトリを決定
    if output_dir is None:
        output_dir = str(original_path.parent)

    # 出力ファイルパスを生成
    output_filename = f"{output_prefix}{original_path.stem}.pdf"
    output_path = Path(output_dir) / output_filename

    temp_pdf = None

    try:
        # Office文書の場合、まずPDFに変換
        if file_ext in OFFICE_EXTENSIONS:
            # 一時ファイルを作成
            temp_dir = tempfile.mkdtemp()
            temp_pdf = Path(temp_dir) / f"{original_path.stem}.pdf"

            success, error_msg = convert_office_to_pdf(file_path, str(temp_pdf))
            if not success:
                # 一時ディレクトリをクリーンアップ
                shutil.rmtree(temp_dir, ignore_errors=True)
                return ProcessResult(
                    success=False,
                    error_message=error_msg,
                    original_filename=original_path.name
                )

            pdf_to_encrypt = str(temp_pdf)
        else:
            pdf_to_encrypt = file_path

        # PDFにパスワードを設定
        success, error_msg = lock_pdf_file(pdf_to_encrypt, str(output_path), password)

        # 一時ファイルをクリーンアップ
        if temp_pdf and temp_pdf.parent.exists():
            shutil.rmtree(temp_pdf.parent, ignore_errors=True)

        if success:
            return ProcessResult(
                success=True,
                output_path=str(output_path),
                original_filename=original_path.name
            )
        else:
            return ProcessResult(
                success=False,
                error_message=error_msg,
                original_filename=original_path.name
            )

    except Exception as e:
        # 一時ファイルをクリーンアップ
        if temp_pdf and temp_pdf.parent.exists():
            shutil.rmtree(temp_pdf.parent, ignore_errors=True)

        return ProcessResult(
            success=False,
            error_message=f"予期しないエラー: {str(e)}",
            original_filename=original_path.name
        )


def process_uploaded_file(
    uploaded_file: BinaryIO,
    filename: str,
    password: str
) -> Tuple[bool, bytes, str]:
    """
    アップロードされたファイルを処理（Webアプリ用）

    Args:
        uploaded_file: アップロードされたファイルオブジェクト
        filename: 元のファイル名
        password: 設定するパスワード

    Returns:
        (成功フラグ, 暗号化されたPDFバイト, エラーメッセージ)
    """
    file_ext = Path(filename).suffix.lower()

    # PDFの場合は直接処理
    if file_ext == '.pdf':
        pdf_bytes = uploaded_file.read()
        return lock_pdf_bytes(pdf_bytes, password)

    # Office文書の場合は一時ファイル経由で変換
    elif file_ext in OFFICE_EXTENSIONS:
        temp_dir = None
        try:
            temp_dir = tempfile.mkdtemp()

            # 入力ファイルを一時保存
            input_temp = Path(temp_dir) / filename
            with open(input_temp, 'wb') as f:
                f.write(uploaded_file.read())

            # PDFに変換
            pdf_temp = Path(temp_dir) / f"{Path(filename).stem}.pdf"
            success, error_msg = convert_office_to_pdf(str(input_temp), str(pdf_temp))

            if not success:
                return False, b"", error_msg

            # 変換されたPDFを読み込んでパスワード設定
            with open(pdf_temp, 'rb') as f:
                pdf_bytes = f.read()

            return lock_pdf_bytes(pdf_bytes, password)

        finally:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)

    else:
        return False, b"", f"未対応のファイル形式です: {file_ext}"


def get_default_output_dir() -> Path:
    """
    デフォルトの出力ディレクトリを取得

    Returns:
        出力ディレクトリのパス（デスクトップの「パスワード付きPDF」フォルダ）
    """
    output_dir = Path.home() / "Desktop" / "パスワード付きPDF"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir
