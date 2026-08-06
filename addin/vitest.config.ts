import { defineConfig } from "vitest/config";
import path from "node:path";

export default defineConfig({
  resolve: {
    alias: { "ledger-engine": path.resolve(__dirname, "../engine/src/index.ts") },
  },
});
