// Writes python/generated/tool_catalog.json from the TypeScript tool catalog.
// Usage: npm run gen:catalog           (write the file)
//        npm run gen:catalog -- --check (exit 1 if the file is out of date)
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { PYTHON_CATALOG_PATH, renderToolCatalog } from "../src/tools/export.js";

const repoRoot = fileURLToPath(new URL("../", import.meta.url));
const target = path.join(repoRoot, PYTHON_CATALOG_PATH);
const expected = renderToolCatalog();

if (process.argv.includes("--check")) {
  let current = "";
  try {
    current = readFileSync(target, "utf8");
  } catch {
    // Missing file: reported below as out of date.
  }
  if (current !== expected) {
    process.stderr.write(`${PYTHON_CATALOG_PATH} is out of date. Run: npm run gen:catalog\n`);
    process.exit(1);
  }
  process.stdout.write(`${PYTHON_CATALOG_PATH} is up to date.\n`);
} else {
  mkdirSync(path.dirname(target), { recursive: true });
  writeFileSync(target, expected);
  process.stdout.write(`Wrote ${PYTHON_CATALOG_PATH}.\n`);
}
