FROM node:22.12-alpine AS builder

# 必要なファイルをコピー
COPY . /app
WORKDIR /app

# 依存関係インストール & ビルド
RUN --mount=type=cache,target=/root/.npm npm install

FROM node:22-alpine AS release

WORKDIR /app

# ビルド成果物と依存ファイルをコピー
COPY --from=builder /app/build /app/build
COPY --from=builder /app/package.json /app/package.json
COPY --from=builder /app/package-lock.json /app/package-lock.json

ENV NODE_ENV=production

RUN npm ci --ignore-scripts --omit-dev

ENTRYPOINT ["node", "build/index.js"]