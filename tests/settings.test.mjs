import assert from 'node:assert/strict';
import { build } from 'esbuild';
import test from 'node:test';

const stub = `
export class PluginSettingTab {
 constructor(app, plugin) { this.app = app; this.plugin = plugin; }
}`;
const bundle = await build({
	entryPoints: ['src/settings.ts'], bundle: true, write: false, platform: 'node', format: 'esm',
	plugins: [{ name: 'settings-host-stub', setup(builder) {
		builder.onResolve({ filter: /^obsidian$/ }, () => ({ path: 'obsidian', namespace: 'stub' }));
		builder.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: stub, loader: 'js' }));
	} }],
});
const { DocxidianSettingTab, DEFAULT_SETTINGS } = await import(`data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].text).toString('base64')}`);
function setup() {
	const plugin = { settings: { ...DEFAULT_SETTINGS }, saves: 0, refreshes: 0,
		async saveSettings() { this.saves++; }, refreshDocxViews() { this.refreshes++; } };
	return { plugin, tab: new DocxidianSettingTab({}, plugin) };
}

test('all plugin settings have searchable declarative controls and current values', () => {
	const { tab } = setup();
	const definitions = tab.getSettingDefinitions();
	assert.deepEqual(definitions.map(d => d.control.key), ['authorName', 'editorLanguage', 'showRuler', 'autosave']);
	for (const definition of definitions) {
		assert(definition.name); assert(definition.desc);
		assert.equal(tab.getControlValue(definition.control.key), DEFAULT_SETTINGS[definition.control.key]);
	}
	assert.equal(Object.keys(definitions[1].control.options).length, 7);
});

test('settings updates preserve validation, persistence and editor refresh behavior', async () => {
	const { plugin, tab } = setup();
	await tab.setControlValue('authorName', '  Tester  ');
	assert.equal(plugin.settings.authorName, 'Tester');
	await tab.setControlValue('authorName', ' ');
	assert.equal(plugin.settings.authorName, 'Obsidian');
	await tab.setControlValue('editorLanguage', 'de');
	assert.equal(plugin.settings.editorLanguage, 'de');
	await tab.setControlValue('editorLanguage', 'invalid');
	assert.equal(plugin.settings.editorLanguage, 'en');
	await tab.setControlValue('showRuler', true); await tab.setControlValue('autosave', false);
	assert.equal(plugin.settings.showRuler, true); assert.equal(plugin.settings.autosave, false);
	assert.equal(plugin.saves, 6); assert.equal(plugin.refreshes, 4);
	await tab.setControlValue('showRuler', 'false'); await tab.setControlValue('unknown', true);
	assert.equal(plugin.saves, 6); assert.equal(plugin.settings.showRuler, true);
});
