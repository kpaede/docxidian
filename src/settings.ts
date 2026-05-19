import { App, PluginSettingTab, Setting } from 'obsidian';
import DocxidianPlugin from './main';

export interface DocxidianSettings {
	authorName: string;
}

export const DEFAULT_SETTINGS: DocxidianSettings = {
	authorName: 'Obsidian',
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
	}
}
