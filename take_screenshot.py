#!/usr/bin/env python3
"""
PDF Lockerのスクリーンショットを撮影するツール

各ステップのスクリーンショットを自動的に撮影して保存します。
Windows環境で実行してください。
"""

import sys
import time
from pathlib import Path

try:
    import tkinter as tk
    from PIL import ImageGrab, Image
except ImportError:
    print("必要なライブラリをインストールしてください:")
    print("pip install pillow")
    sys.exit(1)

# pdf_lockerをインポート
try:
    import pdf_locker
except ImportError:
    print("pdf_locker.pyが見つかりません")
    sys.exit(1)


def take_screenshot(window, filename):
    """ウィンドウのスクリーンショットを撮影"""
    try:
        # ウィンドウの位置とサイズを取得
        x = window.winfo_rootx()
        y = window.winfo_rooty()
        width = window.winfo_width()
        height = window.winfo_height()

        # スクリーンショットを撮影
        screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))

        # 保存
        screenshot_dir = Path("screenshots")
        screenshot_dir.mkdir(exist_ok=True)
        screenshot.save(screenshot_dir / filename)
        print(f"✓ スクリーンショット保存: {filename}")

    except Exception as e:
        print(f"✗ スクリーンショット失敗: {e}")


def main():
    """スクリーンショット撮影のメイン処理"""
    print("=" * 60)
    print("PDF Locker スクリーンショット撮影ツール")
    print("=" * 60)
    print()
    print("このツールは各ステップのスクリーンショットを自動撮影します。")
    print("Windows環境で実行してください。")
    print()
    print("撮影を開始します...")
    print()

    # アプリを起動
    app = pdf_locker.PDFLockerApp()

    # ウィンドウを表示させるために少し待つ
    app.root.update()
    time.sleep(1)

    # ステップ1: ファイル選択画面
    print("ステップ1を撮影中...")
    take_screenshot(app.root, "step1_file_selection.png")

    # ダミーファイルを追加してボタンを有効化
    app.selected_files.append("サンプル.pdf")
    app.file_listbox.insert(tk.END, "📄 サンプル.pdf")
    app.next_btn_step1.config(state=tk.NORMAL)
    app.root.update()
    time.sleep(0.5)
    take_screenshot(app.root, "step1_with_file.png")

    # ステップ2: パスワード入力画面
    print("ステップ2を撮影中...")
    app._show_step(2)
    app.root.update()
    time.sleep(0.5)
    take_screenshot(app.root, "step2_password.png")

    # パスワード入力状態
    app.password_entry.insert(0, "byouin2024")
    app.root.update()
    time.sleep(0.5)
    take_screenshot(app.root, "step2_with_password.png")

    # パスワード表示
    app.show_password_var.set(True)
    app._toggle_password_visibility()
    app.root.update()
    time.sleep(0.5)
    take_screenshot(app.root, "step2_password_visible.png")

    # ステップ3: 完了画面（モックアップ）
    print("ステップ3を撮影中...")
    app._show_step(3)
    app.result_label.config(
        text="✓ 1個のPDFファイルに鍵をかけました\n\n"
             f"保存した場所:\n{Path.home() / 'Desktop' / 'パスワード付きPDF'}\n\n"
             "ファイル名の最初に「鍵付き_」が付いています。"
    )
    app.root.update()
    time.sleep(0.5)
    take_screenshot(app.root, "step3_complete.png")

    print()
    print("=" * 60)
    print("✓ スクリーンショット撮影完了！")
    print(f"保存先: {Path('screenshots').absolute()}")
    print("=" * 60)
    print()
    print("画像ファイル:")
    print("  - step1_file_selection.png    : ステップ1（初期状態）")
    print("  - step1_with_file.png         : ステップ1（ファイル選択後）")
    print("  - step2_password.png          : ステップ2（初期状態）")
    print("  - step2_with_password.png     : ステップ2（パスワード入力後）")
    print("  - step2_password_visible.png  : ステップ2（パスワード表示）")
    print("  - step3_complete.png          : ステップ3（完了画面）")
    print()

    # ウィンドウを閉じる
    app.root.after(3000, app.root.quit)
    app.root.mainloop()


if __name__ == "__main__":
    main()
