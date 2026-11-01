import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export const pause = ms => new Promise(resolve => setTimeout(resolve, ms));

export function createVaultClient(vault) {
	assert(vault, 'Set OBSIDIAN_TEST_VAULT to the vault to test.');
	const binary = process.env.OBSIDIAN_CLI || '/Applications/Obsidian.app/Contents/MacOS/obsidian-cli';
	function cli(...args) {
		const bridge = fileURLToPath(new URL('./cli-response.mjs', import.meta.url));
		const r = spawnSync(process.execPath, [bridge, binary, `vault=${vault}`, ...args], { encoding: 'utf8', timeout: 30000, killSignal: 'SIGKILL', maxBuffer: 30 * 1024 * 1024 });
		assert.equal(r.status, 0, r.stderr || r.stdout || String(r.error));
		assert(!/^Error:/m.test(r.stdout), r.stdout);
		return r.stdout.trim();
	}
	function evaluate(code) {
		const output = cli('eval', `code=${code}`).replace(/^=>\s*/, '');
		try { return JSON.parse(output); } catch { return output; }
	}
	async function waitFor(code, message, timeout = 20000) {
		const end = Date.now() + timeout;
		while (Date.now() < end) {
			if (evaluate(code)) return;
			await pause(150);
		}
		throw new Error(message);
	}
	return { cli, evaluate, waitFor };
}
