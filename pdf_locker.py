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
        self.geometry("350x180")
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

        # パスワード入力
        ttk.Label(frame, text="パスワード:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(frame, show="*", width=30)
        self.password_entry.grid(row=0, column=1, pady=5, padx=5)

        # パスワード確認
        ttk.Label(frame, text="パスワード(確認):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.confirm_entry = ttk.Entry(frame, show="*", width=30)
        self.confirm_entry.grid(row=1, column=1, pady=5, padx=5)

        # ボタンフレーム
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)

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

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Locker - PDFパスワード設定ツール")
        self.root.geometry("600x450")
        self.root.minsize(500, 400)

        # スタイル設定
        self.style = ttk.Style()
        self.style.configure("Title.TLabel", font=("Helvetica", 14, "bold"))
        self.style.configure("Status.TLabel", font=("Helvetica", 10))

        self._create_widgets()
        self._setup_drag_drop()

        # 選択されたファイルリスト
        self.selected_files: List[str] = []

    def _create_widgets(self):
        """メインウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
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
            text="PDFファイルにパスワード保護を追加します（AES-256暗号化）"
        )
        subtitle_label.pack(pady=(0, 10))

        # ファイル選択エリア
        file_frame = ttk.LabelFrame(main_frame, text="PDFファイル", padding="10")
        file_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # ドラッグ&ドロップエリア
        self.drop_frame = ttk.Frame(file_frame, relief="solid", borderwidth=2)
        self.drop_frame.pack(fill=tk.BOTH, expand=True)

        self.drop_label = ttk.Label(
            self.drop_frame,
            text="ここにPDFファイルをドラッグ&ドロップ\nまたは下のボタンでファイルを選択",
            justify=tk.CENTER,
            anchor=tk.CENTER
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH, pady=30)

        # ファイルリスト
        self.file_listbox = tk.Listbox(file_frame, height=6, selectmode=tk.EXTENDED)
        self.file_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        # スクロールバー
        scrollbar = ttk.Scrollbar(self.file_listbox, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.file_listbox.yview)

        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(
            button_frame,
            text="ファイルを選択",
            command=self._select_files,
            width=15
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="クリア",
            command=self._clear_files,
            width=10
        ).pack(side=tk.LEFT, padx=5)

        self.lock_button = ttk.Button(
            button_frame,
            text="パスワードを設定",
            command=self._lock_files,
            width=18
        )
        self.lock_button.pack(side=tk.RIGHT, padx=5)

        # 進捗バー
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ステータスラベル
        self.status_var = tk.StringVar(value="PDFファイルを選択してください")
        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            style="Status.TLabel"
        )
        self.status_label.pack(pady=5)

    def _setup_drag_drop(self):
        """ドラッグ&ドロップの設定（tkinter標準機能）"""
        # tkinterの標準ドラッグ&ドロップは限定的
        # Windows/macOSでの完全なD&Dにはtkinterdnd2が必要だが、
        # exe化時の互換性を考慮してファイル選択ダイアログを主に使用
        pass

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

            self._update_status()

    def _clear_files(self):
        """ファイルリストをクリア"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.progress_var.set(0)
        self._update_status()

    def _update_status(self):
        """ステータスを更新"""
        count = len(self.selected_files)
        if count == 0:
            self.status_var.set("PDFファイルを選択してください")
        else:
            self.status_var.set(f"{count}個のファイルが選択されています")

    def _lock_files(self):
        """選択されたファイルにパスワードを設定"""
        if not self.selected_files:
            messagebox.showwarning("警告", "PDFファイルを選択してください。")
            return

        # パスワード入力ダイアログを表示
        dialog = PasswordDialog(self.root)
        password = dialog.password

        if not password:
            return

        # 保存先フォルダを選択
        save_dir = filedialog.askdirectory(
            title="保存先フォルダを選択（キャンセルで元のフォルダに保存）"
        )

        # 処理開始
        self.lock_button.config(state=tk.DISABLED)
        self.progress_var.set(0)

        # バックグラウンドで処理
        thread = threading.Thread(
            target=self._process_files,
            args=(password, save_dir),
            daemon=True
        )
        thread.start()

    def _process_files(self, password: str, save_dir: Optional[str]):
        """ファイルを処理（バックグラウンドスレッド）"""
        total = len(self.selected_files)
        success_count = 0
        error_files = []

        for i, file_path in enumerate(self.selected_files):
            try:
                self.root.after(0, lambda: self.status_var.set(
                    f"処理中: {Path(file_path).name}"
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

                # 保存先を決定
                original_path = Path(file_path)
                if save_dir:
                    output_path = Path(save_dir) / f"locked_{original_path.name}"
                else:
                    output_path = original_path.parent / f"locked_{original_path.name}"

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

        # 完了処理
        self.root.after(0, lambda: self._on_process_complete(
            success_count, error_files
        ))

    def _on_process_complete(self, success_count: int, error_files: List[tuple]):
        """処理完了時のコールバック"""
        self.lock_button.config(state=tk.NORMAL)

        if error_files:
            error_msg = "\n".join([
                f"・{Path(f).name}: {e}" for f, e in error_files
            ])
            if success_count > 0:
                messagebox.showwarning(
                    "一部完了",
                    f"{success_count}個のファイルにパスワードを設定しました。\n\n"
                    f"エラーが発生したファイル:\n{error_msg}"
                )
            else:
                messagebox.showerror(
                    "エラー",
                    f"すべてのファイルでエラーが発生しました:\n{error_msg}"
                )
        else:
            messagebox.showinfo(
                "完了",
                f"{success_count}個のPDFファイルにパスワードを設定しました！\n\n"
                "ファイル名の先頭に「locked_」が付いて保存されています。"
            )

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
