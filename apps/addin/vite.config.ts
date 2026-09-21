import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import react from "@vitejs/plugin-react";
import devCerts from "office-addin-dev-certs";
import { defineConfig, type Plugin } from "vite";

function manifestPlugin(): Plugin {
  const manifest = resolve("manifest.xml");
  return {
    name: "pair-manifest",
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        if (request.url?.split("?")[0] !== "/manifest.xml") {
          next();
          return;
        }
        void readFile(manifest)
          .then((contents) => {
            response.setHeader("content-type", "application/xml; charset=utf-8");
            response.end(contents);
          })
          .catch(next);
      });
    },
  };
}

export default defineConfig(async ({ command }) => ({
  plugins: [react(), manifestPlugin()],
  build: {
    rollupOptions: { input: { home: resolve("index.html"), taskpane: resolve("taskpane.html") } },
  },
  ...(command === "serve"
    ? {
        server: {
          https: await devCerts.getHttpsServerOptions(),
          host: "localhost",
          port: 3000,
          proxy: {
            "/api": `http://127.0.0.1:${process.env.SOMMELIER_RELAY_PORT ?? "3001"}`,
            "/agent": `http://127.0.0.1:${process.env.SOMMELIER_RELAY_PORT ?? "3001"}`,
            "/connect": {
              target: `ws://127.0.0.1:${process.env.SOMMELIER_RELAY_PORT ?? "3001"}`,
              ws: true,
            },
            "/pair": {
              target: `ws://127.0.0.1:${process.env.SOMMELIER_RELAY_PORT ?? "3001"}`,
              ws: true,
            },
          },
        },
      }
    : {}),
}));
