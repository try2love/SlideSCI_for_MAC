import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [react(), basicSsl()],
  base: "./",
  server: {
    port: 3000,
    proxy: {
      "/native-helper": {
        target: "http://127.0.0.1:17926",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/native-helper/, ""),
      },
    },
  },
  build: {
    outDir: "dist",
    sourcemap: true,
  },
  test: {
    environment: "node",
    include: ["src/tests/**/*.test.ts"],
  },
});
