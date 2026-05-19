import { Notice, Plugin } from 'obsidian';
import { DEFAULT_SETTINGS, DocxidianSettings, DocxidianSettingTab } from './settings';
import { DocxView, VIEW_TYPE_DOCX } from './DocxView';

const DOCX_EXTENSIONS = ['docx'];

export default class DocxidianPlugin extends Plugin {
	settings: DocxidianSettings;

	async onload() {
		await this.loadSettings();

		this.registerView(
			VIEW_TYPE_DOCX,
			(leaf) => new DocxView(leaf, () => this.settings.authorName),
		);
		this.registerExtensions(DOCX_EXTENSIONS, VIEW_TYPE_DOCX);

		this.addCommand({
			id: 'save-current-docx',
			name: 'Save current docx',
			callback: async () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to save it.');
					return;
				}

				await docxView.saveCurrentDocument();
			},
		});

		this.addSettingTab(new DocxidianSettingTab(this.app, this));
	}

	onunload() {
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData() as Partial<DocxidianSettings>);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}
