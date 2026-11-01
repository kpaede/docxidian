import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { createVaultClient, pause } from './obsidian-test-client.mjs';

const vault = process.env.OBSIDIAN_TEST_VAULT;
const { evaluate, waitFor } = createVaultClient(vault);
const results = [];
let initialized = false;
async function check(name, fn) {
	await fn(); results.push({ name, passed: true }); console.log(`PASS ${name}`);
}
try {
	evaluate('(async()=>{if(window.__settingsSuite)throw Error("Settings test already running");const p=app.plugins.plugins.docxidian,tab=app.setting.pluginTabs.find(t=>t.id==="docxidian");if(!tab)throw Error("Settings tab missing");window.__settingsSuite={plugin:p,tab,settings:{...p.settings}};app.setting.open();app.setting.openTabById("docxidian");return true})()');
	initialized = true;
	await waitFor('window.__settingsSuite.tab.containerEl.querySelectorAll(".setting-item").length===4', 'Settings did not render');
	await check('All four declarative settings are registered for Obsidian search', () => {
		const indexed = evaluate('JSON.stringify(window.__settingsSuite.tab.settingItems.map(i=>i.control?.key))');
		assert.deepEqual(indexed, ['authorName', 'editorLanguage', 'showRuler', 'autosave']);
		assert.equal(evaluate('Object.hasOwn(Object.getPrototypeOf(window.__settingsSuite.tab),"display")'), false);
	});
	await check('The native text control validates and persists its value', async () => {
		evaluate('(()=>{const e=window.__settingsSuite.tab.containerEl.querySelector("input[type=text]");e.value="  Beta settings test  ";e.dispatchEvent(new Event("input",{bubbles:true}));return true})()');
		await waitFor('window.__settingsSuite.plugin.settings.authorName==="Beta settings test"', 'Author setting did not update');
		await pause(250);
		assert.equal(evaluate('window.__settingsSuite.plugin.loadData().then(data=>data.authorName)'), 'Beta settings test');
	});
	await check('Native language and toggle controls update and persist settings', async () => {
		evaluate('(()=>{const s=window.__settingsSuite,e=s.tab.containerEl.querySelector("select");s.language=s.settings.editorLanguage==="de"?"en":"de";e.value=s.language;e.dispatchEvent(new Event("change",{bubbles:true}));s.tab.containerEl.querySelectorAll(".checkbox-container")[0].click();return true})()');
		await waitFor('window.__settingsSuite.plugin.settings.editorLanguage===window.__settingsSuite.language&&window.__settingsSuite.plugin.settings.showRuler!==window.__settingsSuite.settings.showRuler', 'Native controls did not update');
		await pause(250);
		const stored = evaluate('window.__settingsSuite.plugin.loadData().then(data=>JSON.stringify({language:data.editorLanguage,ruler:data.showRuler}))');
		assert.equal(stored.language, evaluate('window.__settingsSuite.language'));
		assert.equal(stored.ruler, evaluate('!window.__settingsSuite.settings.showRuler'));
	});
} catch (error) {
	results.push({ passed: false, error: error.stack }); console.error(error); process.exitCode = 1;
} finally {
	if (initialized) {
		evaluate('(async()=>{const s=window.__settingsSuite;Object.assign(s.plugin.settings,s.settings);await s.plugin.saveSettings();s.plugin.refreshDocxViews();app.setting.close();return true})()');
		assert.equal(evaluate('(async()=>{const s=window.__settingsSuite,data=await s.plugin.loadData();return JSON.stringify(data)===JSON.stringify(s.settings)})()'), true, 'Original settings did not restore');
		evaluate('delete window.__settingsSuite');
	}
	if (process.env.OBSIDIAN_TEST_REPORT) writeFileSync(process.env.OBSIDIAN_TEST_REPORT, JSON.stringify({ vault, results }, null, 2));
}
