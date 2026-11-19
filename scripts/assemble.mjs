// SPDX-License-Identifier: MIT OR Apache-2.0
// Gathers the desktop app's pages into dist/: the same files the site
// serves, with the engine beside them (engine/) instead of from a CDN, so
// the installed app works without a network.
import { cpSync, mkdirSync, rmSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const dist = join(root, 'dist');
const engine = process.env.OXI_ENGINE_DIR || join(root, 'node_modules', '@oxidocs', 'wasm');
if (!existsSync(join(engine, 'oxidocs_wasm_bg.wasm'))) {
  throw new Error(`no engine at ${engine}: run npm ci, or set OXI_ENGINE_DIR`);
}
rmSync(dist, { recursive: true, force: true });
mkdirSync(join(dist, 'engine'), { recursive: true });
for (const f of ['index.html', 'app.js', 'model.js', 'blocks.js', 'sample.js', 'template.xlsx', 'assets']) cpSync(join(root, f), join(dist, f), { recursive: true });
for (const f of ['oxidocs_wasm.js', 'oxidocs_wasm_bg.wasm', 'fonts.js', 'LICENSE-MIT', 'LICENSE-APACHE']) cpSync(join(engine, f), join(dist, 'engine', f));
console.log(`assembled ${dist} (engine from ${engine})`);
