# Playwright公式 Ubuntu 22.04 ベースイメージ
FROM mcr.microsoft.com/playwright:v1.58.2-jammy

# 日本語フォント（B2クラウド・e飛伝IIIの文字化け防止）
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /usr/src/app

# 依存関係インストール
COPY package*.json ./
RUN npm ci --omit=dev

# アプリケーションコード
COPY src/ ./src/

# 環境変数（Cloud Run / docker-compose で上書き）
ENV NODE_ENV=production
ENV PORT=8080
ENV HEADLESS=true

EXPOSE 8080

# ヘルスチェック（curlに依存しない方式）
HEALTHCHECK --interval=30s --timeout=10s --retries=3 \
    CMD node -e "fetch('http://localhost:8080/health').then(r=>{if(!r.ok)throw r.status}).catch(()=>process.exit(1))"

CMD [ "node", "src/server.js" ]
