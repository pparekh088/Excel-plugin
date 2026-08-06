/**
 * Export the tool catalogue to shared/agent/ so the SERVER owns it.
 *
 * The catalogue is authoritative instruction content: it tells the model what
 * it may do. It must therefore never be supplied by the client — a browser
 * that can replace the catalogue can grant itself tools. The engine defines
 * it, this script publishes it, and the server loads it from disk at startup.
 *
 * Run: npm run catalogue -w engine  (output is checked in; CI diffs it)
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { TOOLS, UNSUPPORTED_CAPABILITIES, renderToolCatalogue } from "../src/agent/tools";

const outDir = path.resolve(__dirname, "..", "..", "shared", "agent");
fs.mkdirSync(outDir, { recursive: true });

const catalogue = {
  version: 1,
  tools: TOOLS.map((tool) => ({
    name: tool.name,
    category: tool.category,
    access: tool.access,
    risk: tool.risk,
    description: tool.description,
    params: tool.params,
    ...(tool.limitation ? { limitation: tool.limitation } : {}),
  })),
  unsupported: UNSUPPORTED_CAPABILITIES,
};

fs.writeFileSync(
  path.join(outDir, "tool-catalogue.json"),
  JSON.stringify(catalogue, null, 2) + "\n"
);
fs.writeFileSync(path.join(outDir, "tool-catalogue.txt"), renderToolCatalogue() + "\n");

console.log(`wrote ${catalogue.tools.length} tools to shared/agent/`);
