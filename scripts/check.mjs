// ตรวจ syntax ของ Code.gs และสคริปต์ใน JS.html โดยไม่ต้องติดตั้ง dependency
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

const dir = mkdtempSync(join(tmpdir(), 'eco-'));
const checks = [];

const gs = readFileSync('src/Code.gs', 'utf8');
writeFileSync(join(dir, 'code.js'), gs);
checks.push(['src/Code.gs', join(dir, 'code.js')]);

const js = readFileSync('src/JS.html', 'utf8');
const m = js.match(/<script>([\s\S]*)<\/script>/);
if (!m) { console.error('✗ src/JS.html: ไม่พบ <script>'); process.exit(1); }
writeFileSync(join(dir, 'js.js'), m[1]);
checks.push(['src/JS.html', join(dir, 'js.js')]);

const css = readFileSync('src/CSS.html', 'utf8');
if (!/^\s*<style>/.test(css) || !/<\/style>\s*$/.test(css)) { console.error('✗ src/CSS.html: ต้องห่อด้วย <style>…</style>'); process.exit(1); }
console.log('✓ src/CSS.html');

const idx = readFileSync('src/Index.html', 'utf8');
for (const inc of ['CSS', 'JS']) {
  if (!idx.includes(`include('${inc}')`)) { console.error(`✗ src/Index.html ไม่ได้ include('${inc}')`); process.exit(1); }
}
console.log('✓ src/Index.html (include CSS/JS)');

let failed = false;
for (const [label, file] of checks) {
  const r = spawnSync(process.execPath, ['--check', file], { encoding: 'utf8' });
  if (r.status === 0) console.log(`✓ ${label}`);
  else { failed = true; console.error(`✗ ${label}\n${r.stderr}`); }
}
process.exit(failed ? 1 : 0);
