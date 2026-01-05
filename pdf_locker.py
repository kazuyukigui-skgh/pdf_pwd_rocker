#!/usr/bin/env python3
"""
PDF Locker - PDFにパスワードを設定するローカルツール

AES-256暗号化を使用してPDFファイルにパスワード保護を追加します。
Python環境がないPCでも実行できるようにexe化可能です。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional, List
import threading
import subprocess
import platform

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


class PasswordDialog(tk.Toplevel):
    """パスワード入力ダイアログ（確認入力付き）"""

    def __init__(self, parent: tk.Tk, title: str = "パスワード設定"):
        super().__init__(parent)
        self.title(title)
        self.password: Optional[str] = None
        self.resizable(False, False)

        # モーダルダイアログにする
        self.transient(parent)
        self.grab_set()

        # ウィンドウを中央に配置
        self.geometry("350x200")
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        self._create_widgets()

        # Enterキーでパスワード確定
        self.bind("<Return>", lambda e: self._on_ok())
        self.bind("<Escape>", lambda e: self._on_cancel())

        # フォーカスをパスワード入力欄に
        self.password_entry.focus_set()

        # ダイアログが閉じられるまで待機
        self.wait_window()

    def _create_widgets(self):
        """ウィジェットを作成"""
        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        # 説明ラベル
        ttk.Label(
            frame,
            text="PDFを開くときに必要なパスワードを設定します",
            foreground="gray"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 15))

        # パスワード入力
        ttk.Label(frame, text="パスワード:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(frame, show="*", width=30)
        self.password_entry.grid(row=1, column=1, pady=5, padx=5)

        # パスワード確認
        ttk.Label(frame, text="パスワード(確認):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.confirm_entry = ttk.Entry(frame, show="*", width=30)
        self.confirm_entry.grid(row=2, column=1, pady=5, padx=5)

        # ボタンフレーム
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text="OK", command=self._on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self._on_cancel, width=10).pack(side=tk.LEFT, padx=5)

    def _on_ok(self):
        """OKボタン押下時の処理"""
        password = self.password_entry.get()
        confirm = self.confirm_entry.get()

        if not password:
            messagebox.showwarning("警告", "パスワードを入力してください。", parent=self)
            self.password_entry.focus_set()
            return

        if len(password) < 4:
            messagebox.showwarning("警告", "パスワードは4文字以上にしてください。", parent=self)
            self.password_entry.focus_set()
            return

        if password != confirm:
            messagebox.showwarning("警告", "パスワードが一致しません。", parent=self)
            self.confirm_entry.delete(0, tk.END)
            self.confirm_entry.focus_set()
            return

        self.password = password
        self.destroy()

    def _on_cancel(self):
        """キャンセルボタン押下時の処理"""
        self.password = None
        self.destroy()


class PDFLockerApp:
    """PDF Lockerメインアプリケーション"""

    # ステップ定義
    STEP_SELECT = 1
    STEP_PASSWORD = 2
    STEP_COMPLETE = 3

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Locker - PDFパスワード設定ツール")
        self.root.geometry("550x480")
        self.root.minsize(500, 450)

        # スタイル設定
        self.style = ttk.Style()
        self.style.configure("Title.TLabel", font=("Helvetica", 16, "bold"))
        self.style.configure("Step.TLabel", font=("Helvetica", 10))
        self.style.configure("StepActive.TLabel", font=("Helvetica", 10, "bold"))
        self.style.configure("Hint.TLabel", font=("Helvetica", 9), foreground="gray")
        self.style.configure("Big.TButton", font=("Helvetica", 11), padding=10)

        # 現在のステップ
        self.current_step = self.STEP_SELECT

        # 選択されたファイルリスト
        self.selected_files: List[str] = []

        # 最後に保存したフォルダ
        self.last_save_dir: Optional[Path] = None

        self._create_widgets()

    def _create_widgets(self):
        """メインウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # タイトル
        title_label = ttk.Label(
            main_frame,
            text="PDF Locker",
            style="Title.TLabel"
        )
        title_label.pack(pady=(0, 5))

        subtitle_label = ttk.Label(
            main_frame,
            text="PDFファイルにパスワード保護を追加します",
            style="Hint.TLabel"
        )
        subtitle_label.pack(pady=(0, 15))

        # ステップインジケーター
        self._create_step_indicator(main_frame)

        # メインコンテンツエリア
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # ファイル選択エリア
        self._create_file_selection_area(content_frame)

        # アクションエリア
        self._create_action_area(main_frame)

        # 進捗バー
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=(10, 5))

        # ステータスラベル
        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            style="Hint.TLabel"
        )
        self.status_label.pack(pady=5)

        # 初期状態を設定
        self._update_ui_state()

    def _create_step_indicator(self, parent):
        """ステップインジケーターを作成"""
        step_frame = ttk.Frame(parent)
        step_frame.pack(fill=tk.X, pady=(0, 10))

        # 3つのステップを中央揃えで配置
        inner_frame = ttk.Frame(step_frame)
        inner_frame.pack(anchor=tk.CENTER)

        # Step 1
        self.step1_label = ttk.Label(inner_frame, text="① ファイル選択", style="StepActive.TLabel")
        self.step1_label.pack(side=tk.LEFT, padx=10)

        ttk.Label(inner_frame, text="→").pack(side=tk.LEFT, padx=5)

        # Step 2
        self.step2_label = ttk.Label(inner_frame, text="② パスワード設定", style="Step.TLabel")
        self.step2_label.pack(side=tk.LEFT, padx=10)

        ttk.Label(inner_frame, text="→").pack(side=tk.LEFT, padx=5)

        # Step 3
        self.step3_label = ttk.Label(inner_frame, text="③ 完了", style="Step.TLabel")
        self.step3_label.pack(side=tk.LEFT, padx=10)

    def _create_file_selection_area(self, parent):
        """ファイル選択エリアを作成"""
        file_frame = ttk.LabelFrame(parent, text="PDFファイル", padding="10")
        file_frame.pack(fill=tk.BOTH, expand=True)

        # 大きなファイル選択ボタン
        self.select_button = ttk.Button(
            file_frame,
            text="ここをクリックして\nPDFファイルを選択",
            command=self._select_files,
            style="Big.TButton"
        )
        self.select_button.pack(fill=tk.X, pady=10, ipady=15)

        # ヒント
        ttk.Label(
            file_frame,
            text="複数ファイルの選択も可能です（Ctrlキーを押しながらクリック）",
            style="Hint.TLabel"
        ).pack()

        # ファイルリスト
        list_frame = ttk.Frame(file_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.file_listbox = tk.Listbox(list_frame, height=5, selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # ファイル数表示
        self.file_count_var = tk.StringVar(value="選択されたファイル: 0個")
        ttk.Label(
            file_frame,
            textvariable=self.file_count_var,
            style="Hint.TLabel"
        ).pack(anchor=tk.W, pady=(5, 0))

    def _create_action_area(self, parent):
        """アクションエリアを作成"""
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=10)

        # クリアボタン（左側）
        self.clear_button = ttk.Button(
            action_frame,
            text="クリア",
            command=self._clear_files,
            width=12
        )
        self.clear_button.pack(side=tk.LEFT)

        # 実行ボタン（右側・目立つように）
        self.lock_button = ttk.Button(
            action_frame,
            text="次へ：パスワードを設定 →",
            command=self._lock_files,
            style="Big.TButton"
        )
        self.lock_button.pack(side=tk.RIGHT, ipadx=10)

    def _update_step_indicator(self):
        """ステップインジケーターを更新"""
        # リセット
        self.step1_label.configure(style="Step.TLabel")
        self.step2_label.configure(style="Step.TLabel")
        self.step3_label.configure(style="Step.TLabel")

        # 現在のステップをハイライト
        if self.current_step == self.STEP_SELECT:
            self.step1_label.configure(style="StepActive.TLabel")
        elif self.current_step == self.STEP_PASSWORD:
            self.step2_label.configure(style="StepActive.TLabel")
        elif self.current_step == self.STEP_COMPLETE:
            self.step3_label.configure(style="StepActive.TLabel")

    def _update_ui_state(self):
        """UI状態を更新"""
        has_files = len(self.selected_files) > 0

        # ボタンの状態
        if has_files:
            self.lock_button.config(state=tk.NORMAL)
            self.clear_button.config(state=tk.NORMAL)
            self.lock_button.config(text="次へ：パスワードを設定 →")
        else:
            self.lock_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)
            self.lock_button.config(text="ファイルを選択してください")

        # ファイル数表示
        count = len(self.selected_files)
        self.file_count_var.set(f"選択されたファイル: {count}個")

        # ステップインジケーター
        self._update_step_indicator()

    def _select_files(self):
        """ファイル選択ダイアログを開く"""
        files = filedialog.askopenfilenames(
            title="PDFファイルを選択",
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
        )

        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.file_listbox.insert(tk.END, Path(file).name)

            self.current_step = self.STEP_SELECT
            self._update_ui_state()
            self.status_var.set("ファイルを選択しました。「次へ」をクリックしてください。")

    def _clear_files(self):
        """ファイルリストをクリア"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.progress_var.set(0)
        self.current_step = self.STEP_SELECT
        self._update_ui_state()
        self.status_var.set("")

    def _lock_files(self):
        """選択されたファイルにパスワードを設定"""
        if not self.selected_files:
            return

        # ステップ2へ
        self.current_step = self.STEP_PASSWORD
        self._update_step_indicator()

        # パスワード入力ダイアログを表示
        dialog = PasswordDialog(self.root)
        password = dialog.password

        if not password:
            self.current_step = self.STEP_SELECT
            self._update_step_indicator()
            return

        # 処理開始
        self.lock_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.select_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set("処理中...")

        # バックグラウンドで処理
        thread = threading.Thread(
            target=self._process_files,
            args=(password,),
            daemon=True
        )
        thread.start()

    def _process_files(self, password: str):
        """ファイルを処理（バックグラウンドスレッド）"""
        total = len(self.selected_files)
        success_count = 0
        error_files = []
        output_dir = None

        for i, file_path in enumerate(self.selected_files):
            try:
                # ステータス更新
                file_name = Path(file_path).name
                self.root.after(0, lambda fn=file_name: self.status_var.set(
                    f"処理中: {fn}"
                ))

                # PDFを読み込む
                reader = PdfReader(file_path)

                # 既に暗号化されている場合
                if reader.is_encrypted:
                    error_files.append((file_path, "既にパスワードが設定されています"))
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

                # 保存先を決定（元のフォルダに自動保存）
                original_path = Path(file_path)
                output_path = original_path.parent / f"locked_{original_path.name}"
                output_dir = original_path.parent

                # ファイルを保存
                with open(output_path, "wb") as f:
                    writer.write(f)

                success_count += 1

            except PdfReadError as e:
                error_files.append((file_path, f"PDFの読み込みエラー: {str(e)}"))
            except PermissionError:
                error_files.append((file_path, "ファイルへのアクセス権限がありません"))
            except Exception as e:
                error_files.append((file_path, str(e)))

            # 進捗を更新
            progress = ((i + 1) / total) * 100
            self.root.after(0, lambda p=progress: self.progress_var.set(p))

        # 保存先を記録
        self.last_save_dir = output_dir

        # 完了処理
        self.root.after(0, lambda: self._on_process_complete(
            success_count, error_files
        ))

    def _open_folder(self, path: Path):
        """フォルダをエクスプローラーで開く"""
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", path])
            else:  # Linux
                subprocess.run(["xdg-open", path])
        except Exception:
            pass  # フォルダを開けなくても続行

    def _on_process_complete(self, success_count: int, error_files: List[tuple]):
        """処理完了時のコールバック"""
        self.lock_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)
        self.select_button.config(state=tk.NORMAL)

        # ステップ3へ
        self.current_step = self.STEP_COMPLETE
        self._update_step_indicator()

        if error_files:
            error_msg = "\n".join([
                f"・{Path(f).name}: {e}" for f, e in error_files
            ])
            if success_count > 0:
                result = messagebox.askyesno(
                    "一部完了",
                    f"{success_count}個のファイルにパスワードを設定しました。\n"
                    f"ファイル名: locked_元のファイル名.pdf\n\n"
                    f"エラーが発生したファイル:\n{error_msg}\n\n"
                    f"保存先フォルダを開きますか？"
                )
                if result and self.last_save_dir:
                    self._open_folder(self.last_save_dir)
            else:
                messagebox.showerror(
                    "エラー",
                    f"すべてのファイルでエラーが発生しました:\n{error_msg}"
                )
        else:
            result = messagebox.askyesno(
                "完了",
                f"{success_count}個のPDFにパスワードを設定しました！\n\n"
                f"ファイル名: locked_元のファイル名.pdf\n"
                f"保存先: 元のファイルと同じフォルダ\n\n"
                f"保存先フォルダを開きますか？"
            )
            if result and self.last_save_dir:
                self._open_folder(self.last_save_dir)

        self.status_var.set(f"完了: {success_count}個のファイルを処理しました")
        self._clear_files()

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
