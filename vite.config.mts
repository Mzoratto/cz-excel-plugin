import { defineConfig } from "vite";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

export default defineConfig({
  root: ".",
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html")
      }
    }
  },
  server: {
    port: 5173,
    https: false,
    host: "localhost"
  }
});
