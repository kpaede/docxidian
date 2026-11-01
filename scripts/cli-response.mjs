import { spawn } from 'node:child_process';

// Some Obsidian CLI builds keep a process alive after printing its complete
// response. Stop only that CLI process once the response is received; the
// Obsidian app continues running. Test evaluations serialize data on one line.
const child = spawn(process.argv[2], process.argv.slice(3), { stdio: ['ignore', 'pipe', 'pipe'] });
let output = '';
let completed = false;
const timeout = setTimeout(() => finish(1, 'Timed out waiting for an Obsidian CLI response.\n'), 20000);

function finish(status, error = '') {
	if (completed) return;
	completed = true;
	clearTimeout(timeout);
	process.stdout.write(output);
	if (error) process.stderr.write(error);
	child.kill('SIGKILL');
	process.exitCode = status;
}

child.stdout.on('data', data => {
	output += data;
	try {
		JSON.parse(output.replace(/^=>\s*/, ''));
		finish(0);
	} catch {
		if (output.endsWith('\n')) finish(/^Error:/m.test(output) ? 1 : 0);
	}
});
child.stderr.on('data', data => process.stderr.write(data));
child.on('error', error => finish(1, `${error.message}\n`));
child.on('exit', code => finish(code ?? 1));
