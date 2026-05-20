import { App, PluginSettingTab, Setting } from 'obsidian';
import DocxidianPlugin from './main';
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
	autosave: false,
};

export class DocxidianSettingTab extends PluginSettingTab {
	plugin: DocxidianPlugin;

	constructor(app: App, plugin: DocxidianPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		new Setting(containerEl)
			.setName('Author name')
			.setDesc('Used for comments and tracked changes.')
			.addText(text => text
				.setPlaceholder('Obsidian')
				.setValue(this.plugin.settings.authorName)
				.onChange(async (value) => {
					this.plugin.settings.authorName = value.trim() || DEFAULT_SETTINGS.authorName;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Editor language')
			.setDesc('Select the language used by the editor toolbar, dialogs, and messages.')
			.addDropdown(dropdown => {
				for (const option of DOCXIDIAN_LANGUAGE_OPTIONS) {
					dropdown.addOption(option.code, option.label);
				}

				dropdown
					.setValue(this.plugin.settings.editorLanguage)
					.onChange(async (value) => {
						this.plugin.settings.editorLanguage = normalizeDocxidianLanguage(value);
						await this.plugin.saveSettings();
						this.plugin.refreshDocxViews();
					});
			});

		new Setting(containerEl)
			.setName('Show ruler')
			.setDesc('Show the page ruler above the document body.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.showRuler)
				.onChange(async (value) => {
					this.plugin.settings.showRuler = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));

		new Setting(containerEl)
			.setName('Autosave')
			.setDesc('Automatically save the document shortly after changes.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.autosave)
				.onChange(async (value) => {
					this.plugin.settings.autosave = value;
					await this.plugin.saveSettings();
					this.plugin.refreshDocxViews();
				}));
	}
}
