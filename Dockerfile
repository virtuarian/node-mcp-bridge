FROM node:18-alpine

# 必要なパッケージをインストール（ヘルスチェック用）
RUN apk add --no-cache wget

# nodeユーザーを作成・使用（セキュリティ向上）
RUN addgroup -S appgroup && adduser -S appuser -G appgroup
USER appuser

WORKDIR /app

# 依存関係のインストール（キャッシュ最適化）
COPY package*.json ./
RUN npm ci

# ソースコードとtsconfig
COPY --chown=appuser:appgroup tsconfig.json ./
COPY --chown=appuser:appgroup src ./src

# ビルド
RUN npm run build

# 本番環境用の依存関係のみインストール
RUN npm ci --omit=dev

# 必要なディレクトリの作成
# 権限を持つユーザーに戻る必要があります
USER root
RUN mkdir -p config data && chown -R appuser:appgroup config data
USER appuser

# ポートの公開
EXPOSE 3001

# デフォルトの環境変数を設定
ENV NODE_ENV=production \
    PORT=3001 \
    LOG_LEVEL=info

# ヘルスチェックを設定
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
  CMD wget --no-verbose --tries=1 --spider http://localhost:3001/health || exit 1

# アプリケーションの実行
CMD ["node", "dist/index.js"]