#!/usr/bin/env python3
"""
PDF Locker ビルドスクリプト

PyInstallerを使用してexeファイルを生成します。

Usage:
    python build.py          # specファイルを使用してビルド
    python build.py --simple # シンプルなコマンドでビルド
    python build.py --clean  # ビルド成果物をクリーンアップ
"""

import subprocess
import sys
import shutil
from pathlib import Path


def clean_build_artifacts():
    """ビルド成果物をクリーンアップ"""
    dirs_to_remove = ['build', 'dist', '__pycache__']
    files_to_remove = ['*.pyc', '*.pyo']

    project_root = Path(__file__).parent

    for dir_name in dirs_to_remove:
        dir_path = project_root / dir_name
        if dir_path.exists():
            print(f"Removing {dir_path}...")
            shutil.rmtree(dir_path)

    # .pycファイルを削除
    for pattern in files_to_remove:
        for file in project_root.rglob(pattern):
            print(f"Removing {file}...")
            file.unlink()

    print("Clean completed!")


def check_dependencies():
    """依存パッケージをチェック"""
    try:
        import pypdf
        print(f"pypdf version: {pypdf.__version__}")
    except ImportError:
        print("Error: pypdf is not installed.")
        print("Run: pip install pypdf[crypto]")
        return False

    try:
        import tkinterdnd2
        print(f"tkinterdnd2 version: {tkinterdnd2.__version__}")
    except ImportError:
        print("Warning: tkinterdnd2 is not installed.")
        print("Drag & drop will be disabled.")
        print("To enable, run: pip install tkinterdnd2")

    try:
        import PyInstaller
        print(f"PyInstaller version: {PyInstaller.__version__}")
    except ImportError:
        print("Error: PyInstaller is not installed.")
        print("Run: pip install pyinstaller")
        return False

    return True


def build_with_spec():
    """specファイルを使用してビルド"""
    print("Building with spec file...")
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "pdf_locker.spec", "--clean"],
        check=True
    )
    return result.returncode == 0


def build_simple():
    """シンプルなコマンドでビルド"""
    print("Building with simple command...")

    project_root = Path(__file__).parent
    hooks_dir = project_root / "hooks"
    version_file = project_root / "version_info.txt"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", "PDF_Locker",
        "--clean",
        # UPX圧縮を無効化（セキュリティソフトの誤検知を軽減）
        "--noupx",
        # hooksディレクトリを指定
        "--additional-hooks-dir", str(hooks_dir),
        # 不要なモジュールを除外
        "--exclude-module", "matplotlib",
        "--exclude-module", "numpy",
        "--exclude-module", "pandas",
        "--exclude-module", "scipy",
        "--exclude-module", "PIL",
        # tkinterdnd2を含める（データファイルとサブモジュール）
        "--collect-all", "tkinterdnd2",
        "--hidden-import", "tkinterdnd2",
        # pypdfの暗号化関連
        "--hidden-import", "pypdf._crypt_providers",
        "--hidden-import", "pypdf._crypt_providers._cryptography",
        "pdf_locker.py"
    ]

    # Windowsの場合はバージョン情報を追加
    if sys.platform == 'win32' and version_file.exists():
        cmd.insert(-1, "--version-file")
        cmd.insert(-1, str(version_file))

    result = subprocess.run(cmd, check=True)
    return result.returncode == 0


def main():
    """メイン処理"""
    if "--clean" in sys.argv:
        clean_build_artifacts()
        return

    print("=" * 50)
    print("PDF Locker Build Script")
    print("=" * 50)

    # 依存パッケージをチェック
    if not check_dependencies():
        sys.exit(1)

    print()

    # ビルド
    try:
        if "--simple" in sys.argv:
            success = build_simple()
        else:
            success = build_with_spec()

        if success:
            print()
            print("=" * 50)
            print("Build completed successfully!")
            print("=" * 50)
            print()
            print("Output location: dist/PDF_Locker.exe (Windows)")
            print("                 dist/PDF_Locker.app (macOS)")
            print("                 dist/PDF_Locker (Linux)")
            print()
            print("You can copy this file to any PC and run it.")

    except subprocess.CalledProcessError as e:
        print(f"Build failed with error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
