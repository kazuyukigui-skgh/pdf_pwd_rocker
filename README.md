# PDF Locker

PDFファイルにパスワード保護を追加するツールです。
**デスクトップ版**と**Web版**の両方に対応しています。

## 🎯 2つのバージョンから選べます

| バージョン | 用途 | 起動方法 |
|------------|------|----------|
| 🖥️ デスクトップ版 | 個人のPCで使う | `python pdf_locker.py` または `.exe` |
| 🌐 Web版 | 病院内サーバーで複数人で使う | `streamlit run web_app.py` |

---

## 🌐 NEW! Web版（Streamlit）

**ブラウザから使えるWeb版を追加しました！**

病院内のサーバーで動かせば、複数の端末から同時にアクセスできます。

### Web版のメリット
- ✅ インストール不要（ブラウザだけでOK）
- ✅ 複数人で同時利用可能
- ✅ Dockerで簡単デプロイ
- ✅ サーバー1台で管理

### Web版の起動方法

```bash
# ローカルで起動
pip install -r requirements-web.txt
streamlit run web_app.py

# ブラウザで開く
# http://localhost:8501
```

### Dockerで起動（推奨）

```bash
# イメージをビルド
docker build -t pdf-locker-web .

# コンテナを起動
docker run -p 8501:8501 pdf-locker-web

# ブラウザで開く
# http://localhost:8501
```

---

## 🖥️ デスクトップ版（従来版）

このツールは**誰でも安心して使える**ように設計されています。

### 🆕 NEW! Office文書も直接対応

**Word、Excel、PowerPointファイルも直接投げ込めます！**

- 📝 Word文書 (.docx)
- 📊 Excel表 (.xlsx)
- 📽️ PowerPoint資料 (.pptx)
- 📄 PDFファイル (.pdf)

→ すべて自動的にPDFに変換して鍵をかけます

### 主な改善点

- ✅ **3ステップのウィザード形式**：何をすればいいか迷わない
- ✅ **大きなボタンと文字**：読みやすく、押しやすい（16pt～24pt）
- ✅ **保存先は自動**：デスクトップの「パスワード付きPDF」フォルダに自動保存
- ✅ **パスワード表示機能**：確認入力不要、チェックを入れれば見える
- ✅ **優しい日本語**：専門用語なし、分かりやすい言葉だけ
- ✅ **Office文書対応**：Word/Excel/PowerPointも直接処理可能
- ✅ **詳しい手引き書付き**：「使い方ガイド.md」を参照

## 特徴

- **AES-256暗号化**: 業界標準の強力な暗号化方式
- **完全ローカル**: ファイルを外部サーバーに送信しません
- **シンプルなGUI**: 3ステップで誰でも使える
- **Office文書対応**: Word/Excel/PowerPointを自動でPDF化
- **複数ファイル対応**: 一度に複数のファイルを処理可能
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
| デスクトップGUI | tkinter（Python標準） |
| WebUI | Streamlit |
| パッケージング | PyInstaller / Docker |

### アーキテクチャ

```
pdf_pwd_rock/
├── core_logic.py      # 共通ロジック（パスワード設定処理）
├── pdf_locker.py      # デスクトップ版（Tkinter GUI）
├── web_app.py         # Web版（Streamlit）
├── Dockerfile         # Docker用設定
├── requirements.txt   # 全機能用パッケージ
├── requirements-web.txt # Web版用パッケージ（軽量）
└── README.md
```

**ポイント:** `core_logic.py` に共通処理をまとめているので、パスワード設定ルールを変更する場合は1箇所の修正で両方のアプリに反映されます。

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
