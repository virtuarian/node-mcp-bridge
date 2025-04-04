FROM node:18-alpine

# 必要なパッケージをインストール（ヘルスチェック用）
RUN apk add --no-cache wget

# 作業ディレクトリを設定
WORKDIR /app

# 依存関係ファイルをコピー
COPY package*.json ./

# ユーザー作成前に依存関係をインストール（rootで実行）
RUN npm install

# ソースコードとtsconfig
COPY tsconfig.json ./
COPY src ./src

# プロジェクトをビルド
RUN npm run build

# 本番環境用の依存関係のみインストール
RUN npm install --omit=dev

# 非rootユーザーを作成
RUN addgroup -S appgroup && adduser -S appuser -G appgroup

# 必要なディレクトリの作成
RUN mkdir -p config data

# アプリケーションファイルの所有権を変更
RUN chown -R appuser:appgroup /app

# 非rootユーザーに切り替え
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