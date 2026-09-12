import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  optimizeDeps: { exclude: ["@ironcalc/wasm"] },
  server: { proxy: { "/api": "http://localhost:8000" } },
});
