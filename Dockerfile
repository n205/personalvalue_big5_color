FROM python:3.11-slim-bookworm

ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Chrome および Selenium が確実に動作するための依存ライブラリを完全網羅
RUN apt-get update && apt-get install -y \
    wget \
    curl \
    unzip \
    gnupg \
    ca-certificates \
    libnss3 \
    libxss1 \
    libasound2 \
    libxtst6 \
    libgtk-3-0 \
    libgbm1 \
    fonts-liberation \
    libappindicator3-1 \
    xdg-utils \
    poppler-utils \
    --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# Chrome for Testing (137.0.7151.119) 本体のインストール
RUN wget -q -O /tmp/chrome.zip \
    https://storage.googleapis.com/chrome-for-testing-public/137.0.7151.119/linux64/chrome-linux64.zip \
 && unzip /tmp/chrome.zip -d /opt/ \
 && ln -s /opt/chrome-linux64/chrome /usr/local/bin/google-chrome \
 && chmod +x /opt/chrome-linux64/chrome \
 && rm /tmp/chrome.zip

# 対応 ChromeDriver のインストール
RUN wget -q -O /tmp/chromedriver.zip \
    https://storage.googleapis.com/chrome-for-testing-public/137.0.7151.119/linux64/chromedriver-linux64.zip \
 && unzip /tmp/chromedriver.zip -d /opt/ \
 && ln -s /opt/chromedriver-linux64/chromedriver /usr/local/bin/chromedriver \
 && chmod +x /opt/chromedriver-linux64/chromedriver \
 && rm /tmp/chromedriver.zip

# Python ライブラリ関連
COPY requirements.txt requirements.txt
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt

COPY . /app

# タイムアウトを 600秒（10分）に延長して Gunicorn で起動
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "1", "--timeout", "600", "main:app"]
