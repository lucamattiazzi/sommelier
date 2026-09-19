FROM node:22-alpine AS build
WORKDIR /workspace
RUN corepack enable
COPY . .
RUN pnpm install --frozen-lockfile
RUN pnpm bridge:build
RUN pnpm --filter @lucamattiazzi/sommelier-config build
ARG SOMMELIER_PUBLIC_ORIGIN
RUN SOMMELIER_PUBLIC_ORIGIN=${SOMMELIER_PUBLIC_ORIGIN} pnpm --filter @lucamattiazzi/sommelier-addin manifest:production
RUN pnpm --filter @lucamattiazzi/sommelier-addin... build
RUN pnpm --filter @lucamattiazzi/sommelier-server build

FROM node:22-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production HOST=0.0.0.0 PORT=3000 SOMMELIER_STATIC_DIR=/app/addin
ENV SOMMELIER_BRIDGE_SCRIPT=/app/agent/session.mjs
COPY --from=build /workspace/apps/server/dist /app/server
COPY --from=build /workspace/apps/addin/dist /app/addin
COPY --from=build /workspace/skills/sommelier/scripts /app/agent
USER node
EXPOSE 3000
CMD ["node", "/app/server/main.cjs"]
