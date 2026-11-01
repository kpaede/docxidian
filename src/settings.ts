import { App, PluginSettingTab, type SettingDefinitionControl } from 'obsidian';
import type DocxidianPlugin from './main';
import { DOCXIDIAN_LANGUAGE_OPTIONS, DEFAULT_LANGUAGE, normalizeDocxidianLanguage, type DocxidianLanguage } from './locales';

export interface DocxidianSettings {
	authorName: string;
	editorLanguage: DocxidianLanguage;
	showRuler: boolean;
	autosave: boolean;
}

export const DEFAULT_SETTINGS: DocxidianSettings = {
	authorName: 'Obsidian',
	editorLanguage: DEFAULT_LANGUAGE,
	showRuler: false,
	autosave: true,
};

export class DocxidianSettingTab extends PluginSettingTab {
	constructor(app: App, public plugin: DocxidianPlugin) {
		super(app, plugin);
	}

	getSettingDefinitions(): SettingDefinitionControl<keyof DocxidianSettings>[] {
		return [
			{
				name: 'Author name',
				desc: 'Used for comments and tracked changes.',
				control: { type: 'text', key: 'authorName', placeholder: 'Obsidian', defaultValue: DEFAULT_SETTINGS.authorName },
			},
			{
				name: 'Editor language',
				desc: 'Select the language used by the editor toolbar, dialogs, and messages.',
				control: {
					type: 'dropdown', key: 'editorLanguage', defaultValue: DEFAULT_SETTINGS.editorLanguage,
					options: Object.fromEntries(DOCXIDIAN_LANGUAGE_OPTIONS.map(option => [option.code, option.label])),
				},
			},
			{
				name: 'Show ruler',
				desc: 'Show the page ruler above the document body.',
				control: { type: 'toggle', key: 'showRuler', defaultValue: DEFAULT_SETTINGS.showRuler },
			},
			{
				name: 'Autosave',
				desc: 'Automatically save the document shortly after changes.',
				control: { type: 'toggle', key: 'autosave', defaultValue: DEFAULT_SETTINGS.autosave },
			},
		];
	}

	getControlValue(key: string): unknown {
		switch (key) {
			case 'authorName': return this.plugin.settings.authorName;
			case 'editorLanguage': return this.plugin.settings.editorLanguage;
			case 'showRuler': return this.plugin.settings.showRuler;
			case 'autosave': return this.plugin.settings.autosave;
			default: return undefined;
		}
	}

	async setControlValue(key: string, value: unknown): Promise<void> {
		switch (key) {
			case 'authorName':
				if (typeof value !== 'string') return;
				this.plugin.settings.authorName = value.trim() || DEFAULT_SETTINGS.authorName;
				break;
			case 'editorLanguage':
				if (typeof value !== 'string') return;
				this.plugin.settings.editorLanguage = normalizeDocxidianLanguage(value);
				break;
			case 'showRuler':
			case 'autosave':
				if (typeof value !== 'boolean') return;
				this.plugin.settings[key] = value;
				break;
			default: return;
		}
		await this.plugin.saveSettings();
		if (key !== 'authorName') this.plugin.refreshDocxViews();
	}

}
