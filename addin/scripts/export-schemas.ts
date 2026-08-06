/**
 * Export the Zod tool schemas (source of truth, src/tools/schemas.ts) to JSON
 * Schema files in shared/schemas/ for the server to validate against, plus a
 * manifest.json with per-tool metadata (access, risk, description).
 *
 * Run: npm run schemas   (output is checked in; CI re-runs and diffs)
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { z } from "zod";
import { toolDefinitions } from "../src/tools/schemas";

const outDir = path.resolve(__dirname, "..", "..", "shared", "schemas");
fs.mkdirSync(outDir, { recursive: true });

function write(fileName: string, data: unknown): void {
  const target = path.join(outDir, fileName);
  fs.writeFileSync(target, JSON.stringify(data, null, 2) + "\n");
  console.log(`wrote ${path.relative(process.cwd(), target)}`);
}

const manifest: {
  version: number;
  tools: Record<string, { access: string; risk: string; description: string }>;
} = { version: 1, tools: {} };

for (const definition of Object.values(toolDefinitions)) {
  // io:"input" -> the schema accepts pre-parse input (fields with Zod
  // defaults stay optional). The client always parses before sending, so the
  // wire carries post-default params; both ends accept the traffic.
  const params = z.toJSONSchema(definition.params, { target: "draft-2020-12", io: "input" });
  const result = z.toJSONSchema(definition.result, { target: "draft-2020-12" });
  write(`${definition.name}.params.json`, params);
  write(`${definition.name}.result.json`, result);
  manifest.tools[definition.name] = {
    access: definition.access,
    risk: definition.risk,
    description: definition.description,
  };
}

write("manifest.json", manifest);
