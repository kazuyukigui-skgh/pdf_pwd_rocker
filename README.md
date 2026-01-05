# PDF Locker

PDFファイルにパスワード保護を追加するローカルツールです。

## 特徴

- **AES-256暗号化**: 業界標準の強力な暗号化方式
- **完全ローカル**: ファイルを外部サーバーに送信しません
- **シンプルなGUI**: 直感的な操作でファイルを暗号化
- **複数ファイル対応**: 一度に複数のPDFを処理可能
- **exe化対応**: Pythonがないパソコンでも実行可能

## 必要環境

### 開発環境（ビルド用）
- Python 3.8以上
- pypdf[crypto]
- pyinstaller

### 実行環境
- **exe版**: Windows 10/11（Python不要）
- **app版**: macOS 10.14以上（Python不要）

## セットアップ

### 1. 開発環境の準備

```bash
# リポジトリをクローン（または展開）
cd pdf_pwd_rocker

# 仮想環境を作成（推奨）
python -m venv venv

# 仮想環境を有効化
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 依存パッケージをインストール
pip install -r requirements.txt
```

### 2. 動作確認

```bash
# Pythonから直接実行
python pdf_locker.py
```

### 3. exe/appファイルの生成

```bash
# ビルドスクリプトを使用（推奨）
python build.py

# または、シンプルなコマンドで
python build.py --simple

# ビルド成果物のクリーンアップ
python build.py --clean
```

生成されたファイルは `dist/` ディレクトリに作成されます：
- Windows: `dist/PDF_Locker.exe`
- macOS: `dist/PDF_Locker.app`

## 使い方

1. **PDF Lockerを起動**
   - exe版: `PDF_Locker.exe` をダブルクリック
   - Python版: `python pdf_locker.py`

2. **PDFファイルを選択**
   - 「ファイルを選択」ボタンをクリック
   - 複数ファイルの選択も可能

3. **パスワードを設定**
   - 「パスワードを設定」ボタンをクリック
   - パスワードを入力（確認のため2回入力）

4. **保存先を選択**
   - 保存先フォルダを選択（キャンセルで元のフォルダに保存）
   - ファイル名は `locked_元のファイル名.pdf` として保存

## 配布方法

1. `dist/PDF_Locker.exe` をUSBメモリやネットワーク共有でコピー
2. 配布先のPCでダブルクリックして実行

## 注意事項

### ウイルス対策ソフトの誤検知

PyInstallerで作成したexeファイルは、デジタル署名がないため、セキュリティソフトが「不明なファイル」として警告する場合があります。

**対策**:
- セキュリティソフトの除外設定に追加
- 管理者として実行
- 社内のIT部門に事前確認

### ファイルサイズ

exeファイルには、Pythonランタイムとライブラリが含まれるため、20〜40MB程度のサイズになります。

### 起動時間

初回起動時は内部ファイルの展開のため、数秒〜十数秒かかる場合があります。

## 技術仕様

| 項目 | 内容 |
|------|------|
| 暗号化方式 | AES-256 |
| PDFライブラリ | pypdf |
| GUIフレームワーク | tkinter（Python標準） |
| パッケージング | PyInstaller |

## トラブルシューティング

### 「cryptography」関連のエラー

AES-256暗号化には `cryptography` ライブラリが必要です：

```bash
pip install pypdf[crypto]
```

### exeが起動しない

1. ウイルス対策ソフトの除外設定を確認
2. 管理者として実行を試す
3. 別のフォルダにコピーして実行

### パスワード設定済みPDFを処理できない

既にパスワードが設定されているPDFは処理できません。
先にパスワードを解除してから再度お試しください。

## ライセンス

MIT License

## 更新履歴

- v1.0.0 - 初回リリース
  - 基本的なPDFパスワード設定機能
  - 複数ファイル対応
  - AES-256暗号化
