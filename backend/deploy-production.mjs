import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const { WORKERS_CI_BRANCH, ADMIN_EMAIL, ADMIN_SUB } = process.env;
if (WORKERS_CI_BRANCH !== 'main' || !ADMIN_EMAIL?.trim() || !ADMIN_SUB?.trim()) {
  throw new Error('Se requiere la rama main y los secretos ADMIN_EMAIL y ADMIN_SUB en Cloudflare Builds.');
}

const temp = mkdtempSync(join(tmpdir(), 'inventario-deploy-'));
try {
  const file = join(temp, 'secrets.json');
  writeFileSync(file, JSON.stringify({
    ADMIN_EMAIL: ADMIN_EMAIL.trim(),
    ADMIN_SUB: ADMIN_SUB.trim()
  }), { mode: 0o600 });
  const wrangler = fileURLToPath(new URL('./node_modules/wrangler/bin/wrangler.js', import.meta.url));
  const result = spawnSync(process.execPath, [wrangler, 'deploy', '--secrets-file', file], {
    cwd: fileURLToPath(new URL('.', import.meta.url)), stdio: 'inherit'
  });
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  rmSync(temp, { recursive: true, force: true });
}
