#!/usr/bin/env python3
"""
PDF Locker - PDFに鍵をかけるツール（シニア向けシンプル版）

70代の方でも簡単に使えるよう、ウィザード形式で分かりやすく設計されています。
AES-256暗号化を使用してPDFファイルにパスワード保護を追加します。

特徴:
- 3ステップのシンプルな操作
- 大きなボタンと文字
- 保存先は自動（デスクトップの「パスワード付きPDF」フォルダ）
- パスワードの表示機能付き
- 優しい日本語のメッセージ
- Word/Excel/PowerPoint文書も直接対応

対応形式:
- PDF (.pdf)
- Word文書 (.docx)
- Excel表 (.xlsx)
- PowerPoint資料 (.pptx)

Copyright (c) 2024-2026 KazuyukiGui
Licensed under the Elastic License 2.0

商用利用には有料ライセンスが必要です。詳細はLICENSEファイルをご確認ください。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional, List, Tuple
import threading
import tempfile
import shutil


def _setup_tkdnd_path():
    """PyInstallerでバンドルされた場合にtkdndのパスを設定"""
    if getattr(sys, 'frozen', False):
        # PyInstallerでバンドルされた実行ファイルの場合
        bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        # tkinterdnd2のパスを環境変数に追加
        tkdnd_path = os.path.join(bundle_dir, 'tkinterdnd2', 'tkdnd')
        if os.path.exists(tkdnd_path):
            os.environ['TKDND_LIBRARY'] = tkdnd_path
        # 代替パス（Windowsの場合）
        tkdnd_path_alt = os.path.join(bundle_dir, 'tkdnd')
        if os.path.exists(tkdnd_path_alt):
            os.environ['TKDND_LIBRARY'] = tkdnd_path_alt


# PyInstallerの場合、tkdndパスを先に設定
_setup_tkdnd_path()

# tkinterdnd2のインポート（ドラッグ&ドロップ機能）
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
except Exception:
    # その他のエラー（DLLロードエラーなど）
    DND_AVAILABLE = False

# pypdfのインポート
try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.errors import PdfReadError
except ImportError:
    messagebox.showerror(
        "エラー",
        "pypdfライブラリが見つかりません。\n"
        "pip install pypdf[crypto] を実行してください。"
    )
    sys.exit(1)

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


def extract_patient_id(filename: str) -> Optional[str]:
    """
    ファイル名から患者IDを抽出する

    Args:
        filename: ファイル名

    Returns:
        患者ID（7〜10桁の数字）、見つからない場合はNone
    """
    import re

    # 患者IDのパターン（7〜10桁の数字）
    # 例: 12345678, [12345678], 12345678_山田太郎, CT_12345678_20260110
    pattern = r'\b(\d{7,10})\b'

    match = re.search(pattern, filename)
    if match:
        return match.group(1)

    return None


def generate_password(patient_id: str, pattern: str, birth_date: str = "") -> str:
    """
    患者IDとパターンからパスワードを生成する

    Args:
        patient_id: 患者ID
        pattern: パスワードパターン（'id_only', 'birth_only', 'id_mmdd', 'id_yyyymmdd', 'custom'）
        birth_date: 生年月日（YYYYMMDD形式）

    Returns:
        生成されたパスワード
    """
    if pattern == 'id_only':
        return patient_id

    elif pattern == 'birth_only':
        return birth_date

    elif pattern == 'id_mmdd':
        # 患者ID + 月日（MMDD）
        if len(birth_date) >= 8:
            mmdd = birth_date[4:8]  # YYYYMMDD → MMDD
            return f"{patient_id}-{mmdd}"
        return patient_id

    elif pattern == 'id_yyyymmdd':
        # 患者ID + 年月日（YYYYMMDD）
        if len(birth_date) >= 8:
            return f"{patient_id}-{birth_date}"
        return patient_id

    else:  # custom
        return ""


class PDFLockerApp:
    """PDF Lockerメインアプリケーション（シニア向けシンプル版）"""

    def __init__(self):
        # TkinterDnDが利用可能な場合はそちらを使用（ドラッグ&ドロップ対応）
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("PDFに鍵をかけるツール")
        self.root.geometry("700x550")
        self.root.minsize(700, 550)

        # スタイル設定（大きなフォント）
        self.style = ttk.Style()
        self.style.configure("Title.TLabel", font=("Yu Gothic UI", 24, "bold"))
        self.style.configure("Step.TLabel", font=("Yu Gothic UI", 18, "bold"))
        self.style.configure("Instruction.TLabel", font=("Yu Gothic UI", 14))
        self.style.configure("Big.TButton", font=("Yu Gothic UI", 16, "bold"))
        self.style.configure("Status.TLabel", font=("Yu Gothic UI", 12))

        # ウィザードのステップ管理
        self.current_step = 1  # 1: ファイル選択, 2: パスワード入力, 3: 完了
        self.selected_files: List[str] = []
        self.password: str = ""

        # 患者ID自動認識関連
        self.detected_patient_id: Optional[str] = None
        self.password_pattern: tk.StringVar = tk.StringVar(value='id_mmdd')  # デフォルト: 患者ID + MMDD
        self.birth_date_input: str = ""

        self._create_widgets()
        self._show_step(1)

    def _create_widgets(self):
        """メインウィジェットを作成（ウィザード形式）"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # タイトルエリア（常に表示）
        title_label = ttk.Label(
            main_frame,
            text="🔒 PDFに鍵をかけるツール",
            style="Title.TLabel"
        )
        title_label.pack(pady=(0, 20))

        # ステップ表示エリア（常に表示）
        self.step_frame = ttk.Frame(main_frame)
        self.step_frame.pack(fill=tk.X, pady=(0, 20))

        self.step_labels = []
        steps = ["①PDFを選ぶ", "②パスワードを決める", "③完了"]
        for i, step_text in enumerate(steps, 1):
            label = ttk.Label(
                self.step_frame,
                text=step_text,
                font=("Yu Gothic UI", 14),
                relief="solid",
                borderwidth=2,
                padding=10
            )
            label.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            self.step_labels.append(label)

        # コンテンツエリア（ステップごとに切り替わる）
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        # ステップ1: ファイル選択画面
        self.step1_frame = ttk.Frame(self.content_frame)
        self._create_step1_widgets()

        # ステップ2: パスワード入力画面
        self.step2_frame = ttk.Frame(self.content_frame)
        self._create_step2_widgets()

        # ステップ3: 完了画面
        self.step3_frame = ttk.Frame(self.content_frame)
        self._create_step3_widgets()

    def _create_step1_widgets(self):
        """ステップ1: ファイル選択画面"""
        # 説明文
        instruction = ttk.Label(
            self.step1_frame,
            text="鍵をかけたいファイルを選んでください\n（PDF、Word、Excel、PowerPointが使えます）",
            style="Instruction.TLabel",
            justify=tk.CENTER
        )
        instruction.pack(pady=(20, 30))

        # 選択されたファイル表示エリア
        self.file_display_frame = ttk.LabelFrame(
            self.step1_frame,
            text="選んだファイル",
            padding=15
        )
        self.file_display_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        self.file_listbox = tk.Listbox(
            self.file_display_frame,
            height=8,
            font=("Yu Gothic UI", 12),
            selectmode=tk.SINGLE
        )
        self.file_listbox.pack(fill=tk.BOTH, expand=True)

        # ボタンエリア
        button_area = ttk.Frame(self.step1_frame)
        button_area.pack(fill=tk.X, pady=20)

        # ファイル選択ボタン（大きく）
        select_btn = tk.Button(
            button_area,
            text="📁 ファイルを選ぶ",
            command=self._select_files,
            font=("Yu Gothic UI", 18, "bold"),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            relief="raised",
            borderwidth=3,
            cursor="hand2",
            height=2
        )
        select_btn.pack(fill=tk.X, pady=(0, 10))

        # クリアボタンと次へボタン
        bottom_buttons = ttk.Frame(button_area)
        bottom_buttons.pack(fill=tk.X)

        clear_btn = tk.Button(
            bottom_buttons,
            text="クリア（最初から）",
            command=self._clear_files,
            font=("Yu Gothic UI", 12),
            bg="#f44336",
            fg="white",
            activebackground="#da190b",
            cursor="hand2"
        )
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.next_btn_step1 = tk.Button(
            bottom_buttons,
            text="次へ ▶",
            command=lambda: self._show_step(2),
            font=("Yu Gothic UI", 16, "bold"),
            bg="#2196F3",
            fg="white",
            activebackground="#0b7dda",
            cursor="hand2",
            state=tk.DISABLED,
            height=1,
            width=15
        )
        self.next_btn_step1.pack(side=tk.RIGHT)

    def _create_step2_widgets(self):
        """ステップ2: パスワード生成ルール選択画面"""
        # スクロール可能にする
        canvas = tk.Canvas(self.step2_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.step2_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 説明文
        instruction = ttk.Label(
            scrollable_frame,
            text="パスワードの作り方を選んでください",
            style="Instruction.TLabel",
            justify=tk.CENTER
        )
        instruction.pack(pady=(10, 20))

        # 患者ID表示エリア
        patient_id_frame = ttk.LabelFrame(
            scrollable_frame,
            text="📄 選んだファイル",
            padding=15
        )
        patient_id_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        self.patient_id_label = ttk.Label(
            patient_id_frame,
            text="",
            font=("Yu Gothic UI", 12),
            justify=tk.LEFT
        )
        self.patient_id_label.pack(anchor=tk.W)

        # パスワードパターン選択エリア
        pattern_frame = ttk.LabelFrame(
            scrollable_frame,
            text="パスワードの作り方を選ぶ",
            padding=15
        )
        pattern_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        # パターン1: 患者ID + MMDD（推奨）
        pattern1_frame = tk.Frame(pattern_frame, relief="groove", borderwidth=2, bg="#E8F5E9")
        pattern1_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(
            pattern1_frame,
            text="パターン1: 患者ID + 生年月日（月日）【推奨】",
            variable=self.password_pattern,
            value='id_mmdd',
            font=("Yu Gothic UI", 13, "bold"),
            bg="#E8F5E9",
            activebackground="#E8F5E9",
            command=self._update_password_preview
        ).pack(anchor=tk.W, padx=10, pady=5)

        ttk.Label(
            pattern1_frame,
            text="例：12345678-0101\n👍 推奨：セキュリティと使いやすさのバランスが良い",
            font=("Yu Gothic UI", 11),
            foreground="#2E7D32",
            background="#E8F5E9"
        ).pack(anchor=tk.W, padx=30, pady=(0, 10))

        # パターン2: 患者ID + YYYYMMDD
        pattern2_frame = tk.Frame(pattern_frame, relief="groove", borderwidth=2)
        pattern2_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(
            pattern2_frame,
            text="パターン2: 患者ID + 生年月日（年月日）",
            variable=self.password_pattern,
            value='id_yyyymmdd',
            font=("Yu Gothic UI", 12),
            command=self._update_password_preview
        ).pack(anchor=tk.W, padx=10, pady=5)

        ttk.Label(
            pattern2_frame,
            text="例：12345678-19800101\n👍 最も安全　⚠️ 少し長い",
            font=("Yu Gothic UI", 10),
            foreground="#616161"
        ).pack(anchor=tk.W, padx=30, pady=(0, 10))

        # パターン3: 患者IDのみ
        pattern3_frame = tk.Frame(pattern_frame, relief="groove", borderwidth=2, bg="#FFF3E0")
        pattern3_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(
            pattern3_frame,
            text="パターン3: 患者IDのみ",
            variable=self.password_pattern,
            value='id_only',
            font=("Yu Gothic UI", 12),
            bg="#FFF3E0",
            activebackground="#FFF3E0",
            command=self._update_password_preview
        ).pack(anchor=tk.W, padx=10, pady=5)

        ttk.Label(
            pattern3_frame,
            text="例：12345678\n⚠️ セキュリティが弱いため、院内限定で使用してください",
            font=("Yu Gothic UI", 10),
            foreground="#E65100",
            background="#FFF3E0"
        ).pack(anchor=tk.W, padx=30, pady=(0, 10))

        # パターン4: 生年月日のみ
        pattern4_frame = tk.Frame(pattern_frame, relief="groove", borderwidth=2, bg="#FFF3E0")
        pattern4_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(
            pattern4_frame,
            text="パターン4: 生年月日のみ",
            variable=self.password_pattern,
            value='birth_only',
            font=("Yu Gothic UI", 12),
            bg="#FFF3E0",
            activebackground="#FFF3E0",
            command=self._update_password_preview
        ).pack(anchor=tk.W, padx=10, pady=5)

        ttk.Label(
            pattern4_frame,
            text="例：19800101\n⚠️ 患者本人が開く場合のみ推奨",
            font=("Yu Gothic UI", 10),
            foreground="#E65100",
            background="#FFF3E0"
        ).pack(anchor=tk.W, padx=30, pady=(0, 10))

        # パターン5: 自分で決める
        pattern5_frame = tk.Frame(pattern_frame, relief="groove", borderwidth=2)
        pattern5_frame.pack(fill=tk.X, pady=5)

        tk.Radiobutton(
            pattern5_frame,
            text="パターン5: 自分で決める",
            variable=self.password_pattern,
            value='custom',
            font=("Yu Gothic UI", 12),
            command=self._update_password_preview
        ).pack(anchor=tk.W, padx=10, pady=5)

        # 生年月日入力エリア
        birth_date_frame = ttk.LabelFrame(
            scrollable_frame,
            text="生年月日を入力してください（パターン1,2,4で必要）",
            padding=15
        )
        birth_date_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        input_frame = ttk.Frame(birth_date_frame)
        input_frame.pack(fill=tk.X)

        ttk.Label(
            input_frame,
            text="生年月日:",
            font=("Yu Gothic UI", 12)
        ).pack(side=tk.LEFT, padx=(0, 10))

        # 年
        self.birth_year_entry = tk.Entry(
            input_frame,
            font=("Yu Gothic UI", 14),
            width=6
        )
        self.birth_year_entry.pack(side=tk.LEFT, padx=2)
        self.birth_year_entry.insert(0, "1980")
        self.birth_year_entry.bind("<KeyRelease>", lambda e: self._update_password_preview())

        ttk.Label(input_frame, text="年", font=("Yu Gothic UI", 12)).pack(side=tk.LEFT, padx=2)

        # 月
        self.birth_month_entry = tk.Entry(
            input_frame,
            font=("Yu Gothic UI", 14),
            width=4
        )
        self.birth_month_entry.pack(side=tk.LEFT, padx=2)
        self.birth_month_entry.insert(0, "01")
        self.birth_month_entry.bind("<KeyRelease>", lambda e: self._update_password_preview())

        ttk.Label(input_frame, text="月", font=("Yu Gothic UI", 12)).pack(side=tk.LEFT, padx=2)

        # 日
        self.birth_day_entry = tk.Entry(
            input_frame,
            font=("Yu Gothic UI", 14),
            width=4
        )
        self.birth_day_entry.pack(side=tk.LEFT, padx=2)
        self.birth_day_entry.insert(0, "01")
        self.birth_day_entry.bind("<KeyRelease>", lambda e: self._update_password_preview())

        ttk.Label(input_frame, text="日", font=("Yu Gothic UI", 12)).pack(side=tk.LEFT, padx=2)

        # カスタムパスワード入力エリア
        custom_password_frame = ttk.LabelFrame(
            scrollable_frame,
            text="自分で決めたパスワード（パターン5の場合）",
            padding=15
        )
        custom_password_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        self.password_entry = tk.Entry(
            custom_password_frame,
            show="●",
            font=("Yu Gothic UI", 16),
            width=30
        )
        self.password_entry.pack(fill=tk.X, pady=(0, 10))
        self.password_entry.bind("<KeyRelease>", lambda e: self._update_password_preview())

        # パスワード表示チェックボックス
        self.show_password_var = tk.BooleanVar()
        show_password_check = tk.Checkbutton(
            custom_password_frame,
            text="パスワードを表示する",
            variable=self.show_password_var,
            command=self._toggle_password_visibility,
            font=("Yu Gothic UI", 11)
        )
        show_password_check.pack(anchor=tk.W)

        # パスワードプレビューエリア
        preview_frame = ttk.LabelFrame(
            scrollable_frame,
            text="💡 プレビュー",
            padding=15
        )
        preview_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        self.password_preview_label = ttk.Label(
            preview_frame,
            text="",
            font=("Yu Gothic UI", 14, "bold"),
            foreground="#2196F3"
        )
        self.password_preview_label.pack(anchor=tk.W)

        # 注意書き
        note = ttk.Label(
            scrollable_frame,
            text="⚠ パスワードは忘れないようにメモしてください\nパスワードを忘れるとPDFが開けなくなります",
            font=("Yu Gothic UI", 11),
            foreground="red",
            justify=tk.LEFT
        )
        note.pack(padx=20, pady=10)

        # ボタンエリア
        button_area = ttk.Frame(self.step2_frame)
        button_area.pack(side="bottom", fill=tk.X, pady=10, padx=20)

        back_btn = tk.Button(
            button_area,
            text="◀ 戻る",
            command=lambda: self._show_step(1),
            font=("Yu Gothic UI", 12),
            bg="#9E9E9E",
            fg="white",
            activebackground="#757575",
            cursor="hand2",
            width=10
        )
        back_btn.pack(side=tk.LEFT)

        self.finish_btn = tk.Button(
            button_area,
            text="鍵をかける ✓",
            command=self._lock_files,
            font=("Yu Gothic UI", 14, "bold"),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            cursor="hand2",
            height=1,
            width=15
        )
        self.finish_btn.pack(side=tk.RIGHT)

        # 進捗バー（初期は非表示）
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            scrollable_frame,
            variable=self.progress_var,
            maximum=100,
            length=400
        )

        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(
            scrollable_frame,
            textvariable=self.status_var,
            font=("Yu Gothic UI", 12),
            foreground="blue"
        )

    def _create_step3_widgets(self):
        """ステップ3: 完了画面"""
        # 完了アイコンと メッセージ
        success_label = ttk.Label(
            self.step3_frame,
            text="✅",
            font=("Yu Gothic UI", 72)
        )
        success_label.pack(pady=(40, 20))

        message_label = ttk.Label(
            self.step3_frame,
            text="鍵をかけ終わりました！",
            font=("Yu Gothic UI", 20, "bold")
        )
        message_label.pack(pady=(0, 30))

        # 保存先の案内
        info_frame = ttk.LabelFrame(
            self.step3_frame,
            text="保存した場所",
            padding=20
        )
        info_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 30))

        self.result_label = ttk.Label(
            info_frame,
            text="",
            font=("Yu Gothic UI", 14),
            justify=tk.LEFT
        )
        self.result_label.pack(anchor=tk.W)

        # ボタンエリア
        button_area = ttk.Frame(self.step3_frame)
        button_area.pack(fill=tk.X, pady=20)

        open_folder_btn = tk.Button(
            button_area,
            text="📁 保存した場所を開く",
            command=self._open_output_folder,
            font=("Yu Gothic UI", 14, "bold"),
            bg="#2196F3",
            fg="white",
            activebackground="#0b7dda",
            cursor="hand2",
            height=2
        )
        open_folder_btn.pack(fill=tk.X, pady=(0, 10))

        finish_btn = tk.Button(
            button_area,
            text="終了",
            command=self.root.quit,
            font=("Yu Gothic UI", 14),
            bg="#9E9E9E",
            fg="white",
            activebackground="#757575",
            cursor="hand2"
        )
        finish_btn.pack(side=tk.LEFT)

        another_btn = tk.Button(
            button_area,
            text="もう一度やる",
            command=self._restart,
            font=("Yu Gothic UI", 14, "bold"),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            cursor="hand2"
        )
        another_btn.pack(side=tk.RIGHT)

    def _show_step(self, step: int):
        """指定されたステップを表示"""
        # 古いステップを非表示
        self.step1_frame.pack_forget()
        self.step2_frame.pack_forget()
        self.step3_frame.pack_forget()

        # ステップ表示を更新
        for i, label in enumerate(self.step_labels, 1):
            if i == step:
                label.config(background="#4CAF50", foreground="white")
            elif i < step:
                label.config(background="#E0E0E0", foreground="black")
            else:
                label.config(background="white", foreground="black")

        # 新しいステップを表示
        self.current_step = step
        if step == 1:
            self.step1_frame.pack(fill=tk.BOTH, expand=True)
        elif step == 2:
            self.step2_frame.pack(fill=tk.BOTH, expand=True)
            # 患者IDと選択したファイルを表示
            if self.selected_files:
                first_file = Path(self.selected_files[0]).name
                if self.detected_patient_id:
                    self.patient_id_label.config(
                        text=f"ファイル名: {first_file}\n✅ 患者IDを自動検出: {self.detected_patient_id}"
                    )
                else:
                    self.patient_id_label.config(
                        text=f"ファイル名: {first_file}\n⚠️ 患者IDが見つかりませんでした（パターン5を使用してください）"
                    )
            # パスワードプレビューを更新
            self._update_password_preview()
        elif step == 3:
            self.step3_frame.pack(fill=tk.BOTH, expand=True)

    def _toggle_password_visibility(self):
        """パスワードの表示/非表示を切り替え"""
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="●")

    def _update_password_preview(self):
        """パスワードプレビューを更新"""
        pattern = self.password_pattern.get()

        # 生年月日を取得
        try:
            year = self.birth_year_entry.get().strip()
            month = self.birth_month_entry.get().strip().zfill(2)
            day = self.birth_day_entry.get().strip().zfill(2)
            birth_date = f"{year}{month}{day}"
        except:
            birth_date = ""

        # パスワード生成
        if pattern == 'custom':
            # カスタムパスワード
            preview_text = f"パスワード: {self.password_entry.get() or '（未入力）'}"
        else:
            if self.detected_patient_id:
                password = generate_password(self.detected_patient_id, pattern, birth_date)
                if password:
                    preview_text = f"生成されるパスワード: {password}"
                else:
                    preview_text = "生年月日を正しく入力してください"
            else:
                preview_text = "患者IDが見つかりませんでした"

        self.password_preview_label.config(text=preview_text)

    def _open_output_folder(self):
        """出力フォルダを開く"""
        output_dir = Path.home() / "Desktop" / "パスワード付きPDF"
        if output_dir.exists():
            if sys.platform == "win32":
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                os.system(f'open "{output_dir}"')
            else:
                os.system(f'xdg-open "{output_dir}"')

    def _restart(self):
        """最初からやり直す"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.password = ""
        self.password_entry.delete(0, tk.END)
        self.show_password_var.set(False)
        self.progress_var.set(0)
        self.detected_patient_id = None
        self.password_pattern.set('id_mmdd')
        self.next_btn_step1.config(state=tk.DISABLED)
        self._show_step(1)

    def _select_files(self):
        """ファイル選択ダイアログを開く（シンプル版・Office文書対応）"""
        files = filedialog.askopenfilenames(
            title="ファイルを選んでください",
            filetypes=[
                ("対応ファイル", "*.pdf *.docx *.xlsx *.pptx"),
                ("PDFファイル", "*.pdf"),
                ("Word文書", "*.docx"),
                ("Excel表", "*.xlsx"),
                ("PowerPoint資料", "*.pptx"),
                ("すべてのファイル", "*.*")
            ]
        )

        if files:
            # サポートされている拡張子
            supported_extensions = {'.pdf', '.docx', '.xlsx', '.pptx'}
            unsupported_files = []

            for file in files:
                file_ext = Path(file).suffix.lower()

                if file_ext not in supported_extensions:
                    unsupported_files.append(Path(file).name)
                    continue

                if file not in self.selected_files:
                    self.selected_files.append(file)
                    # ファイル名とアイコンを表示
                    display_name = self._get_file_display_name(file)
                    self.file_listbox.insert(tk.END, display_name)

            # 患者IDを自動抽出（最初のファイルから）
            if self.selected_files and not self.detected_patient_id:
                first_file = Path(self.selected_files[0]).name
                patient_id = extract_patient_id(first_file)
                if patient_id:
                    self.detected_patient_id = patient_id

            # 「次へ」ボタンを有効化
            if self.selected_files:
                self.next_btn_step1.config(state=tk.NORMAL)

            # ファイル数をわかりやすく表示
            count = len(self.selected_files)
            if count > 0:
                messagebox.showinfo(
                    "ファイルを選びました",
                    f"{count}個のファイルを選びました。\n\n「次へ」ボタンを押してください。"
                )

            # 非対応ファイルがあった場合は警告
            if unsupported_files:
                messagebox.showwarning(
                    "対応していないファイル",
                    f"以下のファイルは対応していません:\n\n" +
                    "\n".join(unsupported_files[:5]) +
                    (f"\n...他 {len(unsupported_files) - 5} ファイル" if len(unsupported_files) > 5 else "") +
                    "\n\n対応形式: PDF、Word、Excel、PowerPoint"
                )

    def _get_file_display_name(self, file_path: str) -> str:
        """ファイルの表示名を取得（アイコン付き）"""
        file_ext = Path(file_path).suffix.lower()
        file_name = Path(file_path).name

        icon_map = {
            '.pdf': '📄',
            '.docx': '📝',
            '.xlsx': '📊',
            '.pptx': '📽️'
        }

        icon = icon_map.get(file_ext, '📁')
        return f"{icon} {file_name}"

    def _clear_files(self):
        """ファイルリストをクリア"""
        if self.selected_files:
            result = messagebox.askyesno(
                "確認",
                "選んだファイルを全部クリアします。\nよろしいですか？"
            )
            if result:
                self.selected_files.clear()
                self.file_listbox.delete(0, tk.END)
                self.next_btn_step1.config(state=tk.DISABLED)

    def _lock_files(self):
        """選択されたファイルにパスワードを設定（シンプル版・患者ID対応）"""
        # パスワード生成
        pattern = self.password_pattern.get()

        if pattern == 'custom':
            # カスタムパスワード
            password = self.password_entry.get().strip()

            if not password:
                messagebox.showwarning(
                    "入力してください",
                    "パスワードを入力してください。"
                )
                self.password_entry.focus_set()
                return

            if len(password) < 4:
                messagebox.showwarning(
                    "短すぎます",
                    "パスワードは4文字以上にしてください。"
                )
                self.password_entry.focus_set()
                return
        else:
            # 自動生成パスワード
            if not self.detected_patient_id:
                messagebox.showerror(
                    "エラー",
                    "患者IDが見つかりませんでした。\n\nパターン5（自分で決める）を選択してください。"
                )
                return

            # 生年月日を取得
            try:
                year = self.birth_year_entry.get().strip()
                month = self.birth_month_entry.get().strip().zfill(2)
                day = self.birth_day_entry.get().strip().zfill(2)
                birth_date = f"{year}{month}{day}"

                # 生年月日の簡易バリデーション
                if pattern in ['id_mmdd', 'id_yyyymmdd', 'birth_only']:
                    if len(year) != 4 or not year.isdigit():
                        raise ValueError("年は4桁の数字で入力してください")
                    if len(month) != 2 or not month.isdigit() or int(month) < 1 or int(month) > 12:
                        raise ValueError("月は1〜12の数字で入力してください")
                    if len(day) != 2 or not day.isdigit() or int(day) < 1 or int(day) > 31:
                        raise ValueError("日は1〜31の数字で入力してください")
            except ValueError as e:
                messagebox.showwarning(
                    "入力エラー",
                    f"生年月日を正しく入力してください。\n\n{str(e)}"
                )
                return
            except:
                birth_date = ""

            password = generate_password(self.detected_patient_id, pattern, birth_date)

            if not password:
                messagebox.showwarning(
                    "入力してください",
                    "生年月日を正しく入力してください。"
                )
                return

        # セキュリティ警告（患者IDのみまたは生年月日のみの場合）
        if pattern in ['id_only', 'birth_only']:
            warning_result = messagebox.askyesnocancel(
                "⚠️ セキュリティ警告",
                f"選択したパスワード（{pattern}）は安全性が低いため、\n"
                "以下の場合のみ使用してください：\n\n"
                "✅ 院内スタッフ間でのやり取り\n"
                "✅ すぐに削除する一時的なファイル\n"
                "✅ 他のセキュリティ対策と併用\n\n"
                "❌ 推奨しない使用方法：\n"
                "❌ インターネット経由での送信\n"
                "❌ 患者本人への送付\n"
                "❌ 他の医療機関への紹介\n\n"
                "このパスワードで本当によろしいですか？\n\n"
                "「はい」= このまま続ける\n"
                "「いいえ」= パターンを変更する\n"
                "「キャンセル」= 中止する"
            )

            if warning_result is None:  # キャンセル
                return
            elif warning_result is False:  # いいえ
                return

        # 確認メッセージ
        result = messagebox.askyesno(
            "確認",
            f"このパスワードで鍵をかけます:\n\n「{password}」\n\nよろしいですか？\n\n※パスワードは忘れないようにメモしてください"
        )

        if not result:
            return

        # 処理開始
        self.finish_btn.config(state=tk.DISABLED)
        self.progress_bar.pack(fill=tk.X, pady=(20, 5))
        self.status_label.pack(pady=(0, 10))
        self.progress_var.set(0)
        self.status_var.set("処理を始めます...")

        # バックグラウンドで処理
        thread = threading.Thread(
            target=self._process_files,
            args=(password,),
            daemon=True
        )
        thread.start()

    def _process_files(self, password: str):
        """ファイルを処理（バックグラウンドスレッド・シンプル版・Office文書対応）"""
        # 保存先フォルダを作成（デスクトップに固定）
        output_dir = Path.home() / "Desktop" / "パスワード付きPDF"
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "エラー",
                f"保存先フォルダを作成できませんでした。\n\nデスクトップに「パスワード付きPDF」フォルダを作ろうとしましたが失敗しました。"
            ))
            return

        # 一時ファイル用ディレクトリ
        temp_dir = None
        try:
            temp_dir = tempfile.mkdtemp()
        except Exception:
            pass

        total = len(self.selected_files)
        success_count = 0
        error_files = []
        self.output_folder = output_dir  # 完了画面で使用

        for i, file_path in enumerate(self.selected_files):
            pdf_path_to_encrypt = None
            is_temp_pdf = False

            try:
                file_name = Path(file_path).name
                file_ext = Path(file_path).suffix.lower()

                self.root.after(0, lambda name=file_name: self.status_var.set(
                    f"処理中: {name}"
                ))

                # Office文書の場合、まずPDFに変換
                if file_ext in ['.docx', '.xlsx', '.pptx']:
                    if temp_dir is None:
                        error_files.append((file_path, "一時ファイルの作成に失敗しました"))
                        continue

                    # Word/Excel/PowerPointをPDFに変換
                    temp_pdf = Path(temp_dir) / f"{Path(file_path).stem}.pdf"

                    # 変換状況をステータスに表示
                    self.root.after(0, lambda name=file_name: self.status_var.set(
                        f"PDFに変換中: {name}"
                    ))

                    success, error_msg = convert_office_to_pdf(file_path, str(temp_pdf))

                    if not success:
                        error_files.append((file_path, error_msg))
                        continue

                    pdf_path_to_encrypt = str(temp_pdf)
                    is_temp_pdf = True

                    # 変換完了後、暗号化処理に移る
                    self.root.after(0, lambda name=file_name: self.status_var.set(
                        f"鍵をかけています: {name}"
                    ))
                else:
                    # 既にPDFの場合
                    pdf_path_to_encrypt = file_path

                # PDFを読み込む
                reader = PdfReader(pdf_path_to_encrypt)

                # 既に暗号化されている場合
                if reader.is_encrypted:
                    error_files.append((file_path, "すでに鍵がかかっています"))
                    continue

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

                # 保存先を決定（デスクトップの「パスワード付きPDF」フォルダ）
                # 元のファイル名を使用（拡張子はpdfに変更）
                original_path = Path(file_path)
                output_filename = f"鍵付き_{original_path.stem}.pdf"
                output_path = output_dir / output_filename

                # ファイルを保存
                with open(output_path, "wb") as f:
                    writer.write(f)

                success_count += 1

            except PdfReadError:
                error_files.append((file_path, "PDFファイルが壊れているかもしれません"))
            except PermissionError:
                error_files.append((file_path, "このファイルは開けません（使用中の可能性）"))
            except Exception as e:
                error_msg = str(e)
                if "Office" in error_msg or "Excel" in error_msg or "PowerPoint" in error_msg:
                    error_files.append((file_path, "Office文書の処理に失敗しました"))
                else:
                    error_files.append((file_path, "エラーが発生しました"))

            # 進捗を更新
            progress = ((i + 1) / total) * 100
            self.root.after(0, lambda p=progress: self.progress_var.set(p))

        # 一時ディレクトリをクリーンアップ
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

        # 完了処理
        self.root.after(0, lambda: self._on_process_complete(
            success_count, error_files
        ))

    def _on_process_complete(self, success_count: int, error_files: List[tuple]):
        """処理完了時のコールバック（シンプル版）"""
        self.finish_btn.config(state=tk.NORMAL)

        # エラーがあった場合
        if error_files:
            error_msg = "\n".join([
                f"・{Path(f).name}\n  → {e}" for f, e in error_files
            ])
            if success_count > 0:
                messagebox.showwarning(
                    "一部できました",
                    f"{success_count}個のファイルに鍵をかけました。\n\n"
                    f"できなかったファイル:\n{error_msg}\n\n"
                    "完了画面に進みます。"
                )
            else:
                messagebox.showerror(
                    "できませんでした",
                    f"すべてのファイルに鍵をかけられませんでした:\n\n{error_msg}\n\n"
                    "ファイルが開かれていないか確認してください。"
                )
                return

        # 完了画面に情報を設定
        output_dir = Path.home() / "Desktop" / "パスワード付きPDF"
        result_text = f"✓ {success_count}個のPDFファイルに鍵をかけました\n\n"
        result_text += f"保存した場所:\n{output_dir}\n\n"
        result_text += "ファイル名の最初に「鍵付き_」が付いています。"

        if error_files:
            result_text += f"\n\n※ {len(error_files)}個のファイルは処理できませんでした"

        self.result_label.config(text=result_text)

        # ステップ3（完了画面）へ
        self._show_step(3)

    def run(self):
        """アプリケーションを実行"""
        # ウィンドウを中央に配置
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

        self.root.mainloop()


def main():
    """メインエントリーポイント"""
    app = PDFLockerApp()
    app.run()


if __name__ == "__main__":
    main()
