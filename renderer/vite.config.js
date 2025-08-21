import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  root: "renderer",
  base: "./",
  build: { outDir: "../dist/renderer", emptyOutDir: true },
  plugins: [react()]
});
