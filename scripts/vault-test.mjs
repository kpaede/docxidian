import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { createDocxFixture } from './docx-fixtures.mjs';
import { createVaultClient, pause } from './obsidian-test-client.mjs';

// Run only against a vault explicitly selected by the caller. All writes are
// confined to a newly created test folder; settings are changed only in memory.
const vault = process.env.OBSIDIAN_TEST_VAULT;
const { cli, evaluate, waitFor } = createVaultClient(vault);
const folder = `Docxidian integration ${Date.now()}`;
const results = [];
function check(name, fn) {
	return Promise.resolve().then(fn).then(() => {
		results.push({ name, passed: true });
		console.log(`PASS ${name}`);
	});
}
async function open(name) {
	evaluate(`(async()=>{const l=window.__docxSuite.leaf;await l.openFile(app.vault.getAbstractFileByPath(${JSON.stringify(`${folder}/${name}`)}));app.workspace.setActiveLeaf(l,{focus:true});return true})()`);
	await waitFor('!!window.__docxSuite.leaf.view.hostEl?.querySelector(".layout-page")', `No pages for ${name}`);
	await pause(700); // dirty tracking starts 500 ms after mounting
}
function pmText() { return evaluate('JSON.stringify(window.__docxSuite.pm().pmViewDesc.node.textContent)'); }
function focusEnd() {
	assert.equal(evaluate('(()=>{app.workspace.setActiveLeaf(window.__docxSuite.leaf,{focus:true});const e=window.__docxSuite.pm();e.focus();const s=getSelection();s.selectAllChildren(e);s.collapseToEnd();return document.activeElement===e})()'), true);
}
function key(key, modifiers = 0, text) {
	const code = key === ' ' ? 'Space' : key.length === 1 ? `Key${key.toUpperCase()}` : key;
	const virtual = key === ' ' ? 32 : key.length === 1 ? key.toUpperCase().charCodeAt(0) : { Enter: 13, Backspace: 8, ArrowLeft: 37, ArrowRight: 39, Tab: 9 }[key];
	cli('dev:cdp', 'method=Input.dispatchKeyEvent', `params=${JSON.stringify({ type: 'keyDown', key, code, modifiers, windowsVirtualKeyCode: virtual, ...(text === undefined ? {} : { text, unmodifiedText: text }) })}`);
	cli('dev:cdp', 'method=Input.dispatchKeyEvent', `params=${JSON.stringify({ type: 'keyUp', key, code, modifiers, windowsVirtualKeyCode: virtual })}`);
}
function type(text) { for (const character of text) key(character, 0, character); }
async function zip(name) {
	const b64 = evaluate(`(async()=>{const b=await app.vault.readBinary(app.vault.getAbstractFileByPath(${JSON.stringify(`${folder}/${name}`)}));let s='';for(const c of new Uint8Array(b))s+=String.fromCharCode(c);return JSON.stringify(btoa(s))})()`);
	return JSZip.loadAsync(b64, { base64: true, checkCRC32: true });
}
async function xml(name) { return (await zip(name)).file('word/document.xml').async('string'); }
function setInput(index, value) {
	evaluate(`(()=>{const e=window.__docxSuite.leaf.view.hostEl.querySelectorAll('.docxidian-find-dialog input')[${index}];Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(e,${JSON.stringify(value)});e.dispatchEvent(new Event('input',{bubbles:true}));return true})()`);
}
function clickButton(text) {
	assert.equal(evaluate(`(()=>{const b=[...window.__docxSuite.leaf.view.hostEl.querySelectorAll('.docxidian-find-dialog button')].find(b=>b.textContent===${JSON.stringify(text)});if(!b)return false;b.click();return true})()`), true);
}

let initialized = false;
try {
	cli('dev:debug', 'on');
	cli('dev:cdp', 'method=Emulation.setFocusEmulationEnabled', 'params={"enabled":true}');
	evaluate(`(async()=>{if(window.__docxSuite)throw Error('Another test suite is active');const p=app.plugins.plugins.docxidian;window.__docxSuite={originalLeaf:app.workspace.activeLeaf,settings:{...p.settings},folder:${JSON.stringify(folder)},leaf:app.workspace.getLeaf('tab')};window.__docxSuite.pm=()=>[...document.querySelectorAll('.paged-editor__hidden-pm .ProseMirror')].find(e=>e.isContentEditable&&e.textContent.startsWith('BetaStart'));await app.vault.createFolder(${JSON.stringify(folder)});return true})()`);
	initialized = true;
	evaluate('(async()=>{const s=window.__docxSuite;s.digest=async b=>Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256",b)),c=>c.toString(16).padStart(2,"0")).join("");const files=[...new Set(app.workspace.getLeavesOfType("docxidian-docx-view").map(l=>l.view.file).filter(Boolean))];s.originalDocuments=await Promise.all(files.map(async file=>({path:file.path,hash:await s.digest(await app.vault.readBinary(file))})));return true})()');
	for (const [name, options] of [['edit.docx', {}], ['review.docx', { review: true }], ['long.docx', { paragraphs: 50 }]]) {
		const b64 = await createDocxFixture(options);
		evaluate(`(async()=>{await app.vault.createBinary(${JSON.stringify(`${folder}/${name}`)},Uint8Array.from(atob(${JSON.stringify(b64)}),c=>c.charCodeAt(0)).buffer);return true})()`);
	}
	evaluate('(()=>{const p=app.plugins.plugins.docxidian;p.settings.autosave=false;p.settings.editorLanguage="en";p.refreshDocxViews();return true})()');
	await open('edit.docx');
	await check('DOCX opens with text, header and table', () => {
		const text = evaluate('JSON.stringify(window.__docxSuite.leaf.view.hostEl.innerText)');
		for (const expected of ['BetaStart', 'Beta header', 'Cell A', 'Cell B']) assert(text.includes(expected), expected);
	});
	await check('Portal uses pre-wrap and actual keystrokes preserve spaces', () => {
		assert.equal(evaluate('getComputedStyle(window.__docxSuite.pm()).whiteSpace'), 'pre-wrap');
		focusEnd(); type(' hello world  two spaces');
		assert(pmText().endsWith(' hello world  two spaces'), pmText());
	});
	await check('Enter and Unicode input survive editing', () => {
		key('Enter'); type('Grüße café 日本語');
		assert(pmText().includes('Grüße café 日本語'), pmText());
	});
	await check('Manual save updates the actual vault DOCX ZIP', async () => {
		assert.equal(evaluate('window.__docxSuite.leaf.view.saveCurrentDocument()'), true);
		const document = await xml('edit.docx');
		assert(document.includes('hello world  two spaces'));
		assert(document.includes('Grüße café 日本語'));
		assert(document.includes('<w:tbl'));
		assert((await zip('edit.docx')).file('word/header1.xml'));
	});
	await check('Closing and reopening preserves saved content', async () => {
		await open('review.docx'); await open('edit.docx');
		assert(pmText().includes('hello world  two spaces'));
		assert(pmText().includes('Grüße café 日本語'));
	});
	await check('Find and replace operates on the loaded ProseMirror document', async () => {
		evaluate('window.__docxSuite.leaf.view.openFindReplaceDialog()');
		await waitFor('!!window.__docxSuite.leaf.view.hostEl.querySelector(".docxidian-find-dialog")', 'Find panel missing');
		setInput(0, 'world'); setInput(1, 'planet'); await pause(150);
		clickButton('Replace all'); await pause(200);
		assert(pmText().includes('hello planet  two spaces'), pmText());
		evaluate('window.__docxSuite.leaf.view.hostEl.querySelector(".docxidian-find-dialog button[aria-label=Close]").click()');
		assert.equal(evaluate('window.__docxSuite.leaf.view.saveCurrentDocument()'), true);
	});
	await check('Autosave writes typed changes without a save action', async () => {
		evaluate('(()=>{app.plugins.plugins.docxidian.settings.autosave=true;app.plugins.plugins.docxidian.refreshDocxViews();return true})()');
		await pause(700); focusEnd(); type(' autosaved text');
		await waitFor('!window.__docxSuite.leaf.view.isDirty', 'Autosave did not finish', 15000);
		const document = await xml('edit.docx');
		assert(document.includes('autosaved text'));
		assert(document.includes('hello planet  two spaces'), 'Settings refresh reverted saved edits');
	});
	await check('Rename stays inside the test folder and preserves DOCX', async () => {
		evaluate('window.__docxSuite.leaf.view.renameFile("renamed")');
		await waitFor(`window.__docxSuite.leaf.view.file?.path===${JSON.stringify(`${folder}/renamed.docx`)}`, 'Rename failed');
		assert((await xml('renamed.docx')).includes('autosaved text'));
		assert(pmText().includes('hello planet  two spaces'), 'Rename reset the editor');
	});
	await check('Comments and tracked changes survive a save and reopen', async () => {
		await open('review.docx');
		assert.equal(evaluate('window.__docxSuite.leaf.view.reserveReviewSidebar'), true);
		assert(evaluate('JSON.stringify(window.__docxSuite.leaf.view.hostEl.innerText)').includes('Beta review comment'));
		assert.equal(evaluate('window.__docxSuite.leaf.view.saveCurrentDocument()'), true);
		const document = await xml('review.docx');
		assert(/<w:ins\b/.test(document), 'Insertion lost');
		assert(/<w:del\b/.test(document), 'Deletion lost');
		assert(/<w:commentReference\b/.test(document), 'Comment anchor lost');
		assert((await (await zip('review.docx')).file('word/comments.xml').async('string')).includes('Beta review comment'));
	});
	await check('Long DOCX paginates into multiple pages', async () => {
		await open('long.docx');
		await waitFor('window.__docxSuite.leaf.view.hostEl.querySelectorAll(".layout-page").length>1', 'No pagination');
	});
	await check('Native Markdown embed renders read-only pages', async () => {
		const portals = evaluate('document.querySelectorAll(".paged-editor__hidden-pm").length');
		evaluate(`(async()=>{const f=await app.vault.create(${JSON.stringify(`${folder}/embed.md`)},${JSON.stringify(`![[${folder}/renamed.docx]]`)});const l=app.workspace.getLeaf('tab');window.__docxSuite.embedLeaf=l;await l.setViewState({type:'markdown',state:{file:f.path,mode:'preview'}});app.workspace.setActiveLeaf(l,{focus:true});return true})()`);
		await waitFor('!!window.__docxSuite.embedLeaf.view.containerEl.querySelector(".docxidian-embed-pages .layout-page")', 'Embed did not render');
		const content = evaluate('JSON.stringify(window.__docxSuite.embedLeaf.view.containerEl.querySelector(".docxidian-embed-pages").textContent)');
		assert(content.includes('hello planet  two spaces'));
		assert.equal(evaluate('window.__docxSuite.embedLeaf.view.containerEl.querySelectorAll(".docxidian-embed-pages [contenteditable=true]").length'), 0);
		// Obsidian can retain a live-preview widget alongside its reading-mode
		// renderer. Both must be released when the Markdown pane closes.
		assert(evaluate('document.querySelectorAll(".paged-editor__hidden-pm").length') <= portals + 2, 'Duplicate embed renderers');
		evaluate('(()=>{window.__docxSuite.embedLeaf.detach();window.__docxSuite.embedLeaf=null;return true})()');
		await waitFor(`document.querySelectorAll('.paged-editor__hidden-pm').length===${portals}`, 'Closing an embed leaked a hidden editor');
	});
	await check('All seven editor languages render the find panel', async () => {
		await open('renamed.docx');
		for (const [language, title] of [['en', 'Find'], ['de', 'Suchen'], ['pl', 'Znajdź'], ['pt-BR', 'Localizar'], ['tr', 'Bul'], ['he', 'חיפוש'], ['zh-CN', '查找']]) {
			evaluate(`(()=>{app.plugins.plugins.docxidian.settings.editorLanguage=${JSON.stringify(language)};app.plugins.plugins.docxidian.refreshDocxViews();return true})()`);
			await pause(200); evaluate('window.__docxSuite.leaf.view.openFindDialog()'); await pause(150);
			assert.equal(evaluate('JSON.stringify(window.__docxSuite.leaf.view.hostEl.querySelector(".docxidian-find-dialog strong")?.textContent)'), title);
		}
	});
	await check('Unsaved edits survive language and ruler changes', async () => {
		evaluate('(()=>{const p=app.plugins.plugins.docxidian;p.settings.editorLanguage="en";p.settings.autosave=false;p.refreshDocxViews();return true})()');
		await pause(700);
		evaluate('window.__docxSuite.leaf.view.hostEl.querySelector(".docxidian-find-dialog button[aria-label=Close]").click()');
		focusEnd(); type(' unsaved marker');
		evaluate('(()=>{const p=app.plugins.plugins.docxidian;p.settings.showRuler=!p.settings.showRuler;p.settings.editorLanguage="de";p.refreshDocxViews();return true})()');
		await pause(400);
		assert(pmText().includes('unsaved marker'));
		assert(pmText().includes('hello planet  two spaces'));
		assert.equal(evaluate('window.__docxSuite.leaf.view.isDirty'), true);
		assert.equal(evaluate('window.__docxSuite.leaf.view.saveCurrentDocument()'), true);
	});
	await check('A failed vault write keeps edits dirty and can be retried', async () => {
		evaluate('(()=>{const p=app.plugins.plugins.docxidian;p.settings.editorLanguage="en";p.settings.autosave=false;p.refreshDocxViews();return true})()');
		await pause(700);
		focusEnd(); type(' retry marker');
		evaluate('(()=>{const s=window.__docxSuite;s.modifyBinary=app.vault.modifyBinary;app.vault.modifyBinary=async function(file,...args){if(file.path===s.leaf.view.file.path)throw Error("Simulated beta write failure");return s.modifyBinary.call(this,file,...args)};return true})()');
		try {
			assert.equal(evaluate('window.__docxSuite.leaf.view.saveCurrentDocument()'), false);
			assert.equal(evaluate('window.__docxSuite.leaf.view.isDirty'), true);
			assert(!(await xml('renamed.docx')).includes('retry marker'));
		} finally {
			evaluate('(()=>{app.vault.modifyBinary=window.__docxSuite.modifyBinary;delete window.__docxSuite.modifyBinary;return true})()');
		}
		await pause(400);
		assert.equal(evaluate('window.__docxSuite.leaf.view.saveCurrentDocument()'), true);
		assert((await xml('renamed.docx')).includes('retry marker'));
	});
	await check('Existing open DOCX files remain byte-for-byte unchanged', async () => {
		assert.equal(evaluate('(async()=>{const s=window.__docxSuite;for(const f of s.originalDocuments)if(await s.digest(await app.vault.readBinary(app.vault.getAbstractFileByPath(f.path)))!==f.hash)return false;return true})()'), true);
	});
} catch (error) {
	results.push({ passed: false, error: error.stack });
	console.error(error);
	process.exitCode = 1;
} finally {
	if (initialized) {
		// Detach only the test panes; never close a user's existing document.
		evaluate('(async()=>{const s=window.__docxSuite;if(s.modifyBinary)app.vault.modifyBinary=s.modifyBinary;if(s.leaf.view.file?.path.startsWith(s.folder+"/"))s.leaf.view.isDirty=false;Object.assign(app.plugins.plugins.docxidian.settings,s.settings);app.plugins.plugins.docxidian.refreshDocxViews();s.embedLeaf?.detach();await s.leaf.detach();app.workspace.setActiveLeaf(s.originalLeaf,{focus:true});delete window.__docxSuite;return true})()');
	}
	cli('dev:cdp', 'method=Emulation.setFocusEmulationEnabled', 'params={"enabled":false}');
	const report = { vault, folder, results };
	if (process.env.OBSIDIAN_TEST_REPORT) writeFileSync(process.env.OBSIDIAN_TEST_REPORT, JSON.stringify(report, null, 2));
	console.log(`Test fixtures retained in: ${folder}`);
}
