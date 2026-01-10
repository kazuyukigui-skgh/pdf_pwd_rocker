# PDF Locker - Webアプリ用Dockerfile
#
# ビルド方法:
#   docker build -t pdf-locker-web .
#
# 起動方法:
#   docker run -p 8501:8501 pdf-locker-web
#
# ブラウザでアクセス:
#   http://localhost:8501

# 軽量なPython環境を使用
FROM python:3.11-slim

# 作業ディレクトリを設定
WORKDIR /app

# 必要なシステムパッケージをインストール
# （PDFライブラリの依存関係に必要）
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Pythonパッケージをインストール
# （Webアプリに必要な最小限のパッケージのみ）
COPY requirements-web.txt .
RUN pip install --no-cache-dir -r requirements-web.txt

# アプリケーションコードをコピー
COPY core_logic.py .
COPY web_app.py .

# Streamlitの設定
# - ブラウザを自動で開かない
# - すべてのIPアドレスからのアクセスを許可
ENV STREAMLIT_SERVER_PORT=8501
ENV STREAMLIT_SERVER_ADDRESS=0.0.0.0
ENV STREAMLIT_SERVER_HEADLESS=true
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# ポートを公開
EXPOSE 8501

# ヘルスチェック
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8501/_stcore/health || exit 1

# アプリケーションを起動
CMD ["streamlit", "run", "web_app.py"]
