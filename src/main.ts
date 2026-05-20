import { Notice, Plugin } from 'obsidian';
import { DEFAULT_SETTINGS, DocxidianSettings, DocxidianSettingTab } from './settings';
import { processDocxEmbeds, registerDocxFileEmbed } from './DocxEmbed';
import { DocxView, VIEW_TYPE_DOCX } from './DocxView';
import { getDocxEditorLocale, normalizeDocxidianLanguage } from './locales';

const DOCX_EXTENSIONS = ['docx'];

export default class DocxidianPlugin extends Plugin {
	settings: DocxidianSettings;

	async onload() {
		await this.loadSettings();

		this.registerView(
			VIEW_TYPE_DOCX,
			(leaf) => new DocxView(
				leaf,
				() => this.settings.authorName,
				() => this.settings.editorLanguage,
				() => getDocxEditorLocale(this.settings.editorLanguage),
				() => this.settings.showRuler,
				() => this.settings.autosave,
			),
		);
		this.registerExtensions(DOCX_EXTENSIONS, VIEW_TYPE_DOCX);
		registerDocxFileEmbed(this, () => getDocxEditorLocale(this.settings.editorLanguage));
		this.registerMarkdownPostProcessor((el, ctx) => {
			processDocxEmbeds(this.app, el, ctx, () => getDocxEditorLocale(this.settings.editorLanguage));
		}, 1000);

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
		this.addCommand({
			id: 'find-in-current-docx',
			name: 'Find in current docx',
			callback: () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to search it.');
					return;
				}

				docxView.openFindDialog();
			},
		});
		this.addCommand({
			id: 'find-replace-in-current-docx',
			name: 'Find and replace in current docx',
			callback: () => {
				const docxView = this.app.workspace.getActiveViewOfType(DocxView);
				if (!docxView) {
					new Notice('Open a docx file to search it.');
					return;
				}

				docxView.openFindReplaceDialog();
			},
		});

		this.addSettingTab(new DocxidianSettingTab(this.app, this));
	}

	onunload() {
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData() as Partial<DocxidianSettings>);
		this.settings.editorLanguage = normalizeDocxidianLanguage(this.settings.editorLanguage);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	refreshDocxViews() {
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_DOCX)) {
			const view = leaf.view;
			if (view instanceof DocxView) {
				view.refreshSettings();
			}
		}
	}
}
