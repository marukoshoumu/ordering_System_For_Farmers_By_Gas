# 1. Playwright公式のFocal(Ubuntu 20.04/22.04)ベースイメージを使用
FROM mcr.microsoft.com/playwright:v1.49.0-jammy

# 2. 日本語フォント(Noto Sans CJK JP)のインストール
# B2クラウド等で日本語が文字化けしないように必須
RUN apt-get update && apt-get install -y \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# 3. アプリケーションディレクトリの作成
WORKDIR /usr/src/app

# 4. 依存関係のコピーとインストール
COPY package*.json ./
RUN npm ci --only=production

# 5. アプリケーションコードのコピー
COPY . .

# 6. 環境変数の設定 (Cloud Runで上書き可能)
ENV NODE_ENV=production
ENV PORT=8080

# 7. ポートの公開
EXPOSE 8080

# 8. ヘルスチェックの設定 (オプションだが推奨)
HEALTHCHECK --interval=30s --timeout=10s --retries=3 \
    CMD curl -f http://localhost:8080/health || exit 1

# 9. アプリケーションの起動
CMD [ "npm", "start" ]
