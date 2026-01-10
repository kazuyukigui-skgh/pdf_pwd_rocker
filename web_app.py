#!/usr/bin/env python3
"""
PDF Locker - Webアプリ版（Streamlit）

ブラウザからPDFにパスワードをかけられるWebアプリです。
病院内のサーバーで動かして、複数の端末から利用できます。

使い方:
    streamlit run web_app.py

Docker環境での起動:
    docker run -p 8501:8501 pdf-locker-web
"""

import streamlit as st
from pathlib import Path

# 共通ロジックをインポート
from core_logic import (
    check_dependencies,
    is_supported_file,
    get_file_type_icon,
    validate_password,
    lock_pdf_bytes,
    process_uploaded_file,
    SUPPORTED_EXTENSIONS,
    PYPDF_AVAILABLE
)


def main():
    """メインアプリケーション"""

    # ページ設定
    st.set_page_config(
        page_title="PDFに鍵をかけるツール",
        page_icon="🔒",
        layout="centered",
        initial_sidebar_state="collapsed"
    )

    # カスタムCSS（シニアフレンドリーな大きな文字）
    st.markdown("""
        <style>
        .main-title {
            font-size: 2.5rem !important;
            font-weight: bold;
            text-align: center;
            margin-bottom: 2rem;
        }
        .step-header {
            font-size: 1.5rem !important;
            font-weight: bold;
            color: #1f77b4;
            margin-top: 1.5rem;
            margin-bottom: 1rem;
        }
        .stButton > button {
            font-size: 1.2rem !important;
            padding: 0.75rem 2rem !important;
            width: 100%;
        }
        .success-box {
            background-color: #d4edda;
            border: 2px solid #28a745;
            border-radius: 10px;
            padding: 1.5rem;
            margin: 1rem 0;
        }
        .warning-box {
            background-color: #fff3cd;
            border: 2px solid #ffc107;
            border-radius: 10px;
            padding: 1rem;
            margin: 1rem 0;
        }
        .file-info {
            background-color: #e7f3ff;
            border-radius: 8px;
            padding: 1rem;
            margin: 0.5rem 0;
        }
        </style>
    """, unsafe_allow_html=True)

    # 依存関係チェック
    deps_ok, deps_error = check_dependencies()
    if not deps_ok:
        st.error(deps_error)
        st.stop()

    # タイトル
    st.markdown('<h1 class="main-title">🔒 PDFに鍵をかけるツール</h1>', unsafe_allow_html=True)

    # 説明文
    st.info("""
    **使い方は簡単です:**
    1. ファイルをアップロード
    2. パスワードを入力
    3. ダウンロード

    対応形式: PDF、Word、Excel、PowerPoint
    """)

    # ステップ1: ファイルアップロード
    st.markdown('<p class="step-header">① ファイルをアップロード</p>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "鍵をかけたいファイルを選んでください",
        type=["pdf", "docx", "xlsx", "pptx"],
        help="PDF、Word(.docx)、Excel(.xlsx)、PowerPoint(.pptx)に対応しています"
    )

    if uploaded_file is not None:
        # ファイル情報を表示
        file_ext = Path(uploaded_file.name).suffix.lower()
        icon = get_file_type_icon(uploaded_file.name)

        st.markdown(f"""
            <div class="file-info">
                <strong>{icon} 選択されたファイル:</strong><br>
                {uploaded_file.name}<br>
                <small>サイズ: {uploaded_file.size / 1024:.1f} KB</small>
            </div>
        """, unsafe_allow_html=True)

        # Office文書の場合の注意
        if file_ext in ['.docx', '.xlsx', '.pptx']:
            st.info("📝 このファイルはPDFに変換してから鍵をかけます。")

        # ステップ2: パスワード入力
        st.markdown('<p class="step-header">② パスワードを決める</p>', unsafe_allow_html=True)

        col1, col2 = st.columns([3, 1])

        with col1:
            password = st.text_input(
                "パスワード（4文字以上）",
                type="password",
                help="PDFを開くときに必要なパスワードです",
                max_chars=50
            )

        with col2:
            show_password = st.checkbox("表示する")

        # パスワード表示
        if show_password and password:
            st.text(f"入力中のパスワード: {password}")

        # パスワードの強さインジケーター
        if password:
            if len(password) < 4:
                st.warning("⚠️ パスワードは4文字以上にしてください")
            elif len(password) < 8:
                st.info("💡 もう少し長いパスワードがおすすめです")
            else:
                st.success("✅ 良いパスワードです")

        # 注意書き
        st.markdown("""
            <div class="warning-box">
                ⚠️ <strong>重要:</strong> パスワードは忘れないようにメモしてください。<br>
                パスワードを忘れるとPDFが開けなくなります。
            </div>
        """, unsafe_allow_html=True)

        # ステップ3: 処理実行
        st.markdown('<p class="step-header">③ 鍵をかける</p>', unsafe_allow_html=True)

        # パスワードの検証
        is_valid, validation_error = validate_password(password)

        if st.button("🔒 鍵をかけてダウンロード", type="primary", disabled=not is_valid):
            with st.spinner("処理中です...しばらくお待ちください"):
                # ファイルを処理
                uploaded_file.seek(0)  # ファイルポインタをリセット

                success, locked_pdf_bytes, error_msg = process_uploaded_file(
                    uploaded_file,
                    uploaded_file.name,
                    password
                )

                if success:
                    # 成功メッセージ
                    st.markdown("""
                        <div class="success-box">
                            <h3>✅ 鍵をかけ終わりました！</h3>
                            <p>下のボタンからダウンロードしてください。</p>
                        </div>
                    """, unsafe_allow_html=True)

                    # ダウンロードボタン
                    output_filename = f"鍵付き_{Path(uploaded_file.name).stem}.pdf"

                    st.download_button(
                        label="📥 ダウンロード",
                        data=locked_pdf_bytes,
                        file_name=output_filename,
                        mime="application/pdf",
                        type="primary"
                    )

                    st.info(f"💡 ダウンロードされるファイル名: **{output_filename}**")

                else:
                    # エラーメッセージ
                    st.error(f"❌ 処理できませんでした\n\n{error_msg}")

    else:
        # ファイルが選択されていない場合
        st.markdown("""
            ---
            ### 📌 対応ファイル形式

            | 形式 | 拡張子 | 説明 |
            |------|--------|------|
            | 📄 PDF | .pdf | そのまま鍵をかけます |
            | 📝 Word | .docx | PDFに変換してから鍵をかけます |
            | 📊 Excel | .xlsx | PDFに変換してから鍵をかけます |
            | 📽️ PowerPoint | .pptx | PDFに変換してから鍵をかけます |

            ---
            ### 🔐 セキュリティについて

            - **AES-256暗号化**: 業界標準の強力な暗号化方式を使用
            - **ローカル処理**: ファイルは外部サーバーに送信されません
            - **安心設計**: 病院での利用を想定した設計です
        """)

    # フッター
    st.markdown("---")
    st.markdown("""
        <div style="text-align: center; color: #888; font-size: 0.9rem;">
            🔒 PDFに鍵をかけるツール | AES-256暗号化対応<br>
            <small>ファイルは外部に送信されません</small>
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
