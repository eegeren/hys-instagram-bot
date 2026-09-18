import { rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

// A fixed output directory rooted at this script's project; never accept an external cleanup path.
await rm(fileURLToPath(new URL('../dist/', import.meta.url)), { recursive: true, force: true });
const result = spawnSync(process.execPath, [fileURLToPath(new URL('../node_modules/typescript/bin/tsc', import.meta.url))], { stdio: 'inherit' });
process.exitCode = result.status ?? 1;
