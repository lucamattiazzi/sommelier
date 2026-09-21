import { defineConfig } from "@lucamattiazzi/sommelier-config/config";

export default defineConfig({
  app: {
    id: "4dc1f4b5-e74d-4b9a-87d1-d36d78a2bd4f",
    name: "Sommelier",
    description: "Excels at pairing. Connect Excel to your AI agent and approve workbook changes.",
    version: "0.2.1",
    providerName: "Luca Mattiazzi",
  },
  office: {
    hosts: ["Workbook"],
    permissions: "ReadWriteDocument",
    requirements: { ExcelApi: "1.13" },
  },
  taskpane: {
    developmentUrl: "https://localhost:3000",
    productionUrl:
      process.env.SOMMELIER_PUBLIC_ORIGIN ??
      process.env.PAIR_PUBLIC_ORIGIN ??
      "https://sommelier.example.com",
    entryPath: "/taskpane.html",
    developmentCommand: "pnpm dev",
    supportPath: "/support.html",
  },
  commands: {
    label: "Open Sommelier",
    groupLabel: "Sommelier",
    icon: "./public/icon.png",
    icons: {
      16: "./public/icon-16.png",
      32: "./public/icon-32.png",
      64: "./public/icon-64.png",
      80: "./public/icon-80.png",
    },
  },
});
