import { de, he, pl, ptBR, tr, zhCN, type Translations } from '@eigenpal/docx-editor-i18n';

export type DocxidianLanguage = 'en' | 'de' | 'pl' | 'pt-BR' | 'tr' | 'he' | 'zh-CN';

export interface DocxidianLanguageOption {
	code: DocxidianLanguage;
	label: string;
}

export const DEFAULT_LANGUAGE: DocxidianLanguage = 'en';

export const DOCXIDIAN_LANGUAGE_OPTIONS: DocxidianLanguageOption[] = [
	{ code: 'en', label: 'English' },
	{ code: 'de', label: 'Deutsch' },
	{ code: 'pl', label: 'Polski' },
	{ code: 'pt-BR', label: 'Portugues do Brasil' },
	{ code: 'tr', label: 'Turkce' },
	{ code: 'he', label: 'Hebrew' },
	{ code: 'zh-CN', label: 'Simplified Chinese' },
];

const DOCX_EDITOR_LOCALES: Record<DocxidianLanguage, Translations | undefined> = {
	en: undefined,
	de,
	pl,
	'pt-BR': ptBR,
	tr,
	he,
	'zh-CN': zhCN,
};

export function isDocxidianLanguage(value: string): value is DocxidianLanguage {
	return Object.prototype.hasOwnProperty.call(DOCX_EDITOR_LOCALES, value);
}

export function normalizeDocxidianLanguage(value: string | undefined): DocxidianLanguage {
	return value && isDocxidianLanguage(value) ? value : DEFAULT_LANGUAGE;
}

export function getDocxEditorLocale(language: DocxidianLanguage): Translations | undefined {
	return DOCX_EDITOR_LOCALES[language];
}
