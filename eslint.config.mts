import obsidianmd from 'eslint-plugin-obsidianmd';
import globals from 'globals';
import tseslint from 'typescript-eslint';
import { globalIgnores } from 'eslint/config';
import { fileURLToPath } from 'node:url';

export default [
	globalIgnores(['node_modules', 'dist', 'main.js', 'package-lock.json']),
	{
		files: ['src/**/*.{ts,tsx,cjs}'],
		languageOptions: { globals: { ...globals.browser, ...globals.commonjs } },
		plugins: { obsidianmd },
		rules: {
			'no-alert': 'error',
			'no-debugger': 'error',
			'no-var': 'error',
			'prefer-const': 'error',
			'obsidianmd/no-tfile-tfolder-cast': 'error',
			'obsidianmd/regex-lookbehind': 'error',
			'obsidianmd/vault/iterate': 'error',
			'obsidianmd/no-static-styles-assignment': 'error',
			'obsidianmd/prefer-window-timers': 'error',
			'obsidianmd/no-global-this': 'error',
			'obsidianmd/settings-tab/prefer-setting-definitions': 'error',
			'obsidianmd/settings-tab/require-display': 'error',
			'obsidianmd/settings-tab/no-deprecated-display': 'error',
		},
	},
	{
		files: ['src/**/*.{ts,tsx}'],
		languageOptions: {
			parser: tseslint.parser,
			parserOptions: { project: './tsconfig.json', tsconfigRootDir: fileURLToPath(new URL('.', import.meta.url)) },
		},
		plugins: { '@typescript-eslint': tseslint.plugin },
		rules: {
			'obsidianmd/prefer-instanceof': 'error',
			'@typescript-eslint/no-unsafe-return': 'error',
			'@typescript-eslint/no-misused-promises': 'error',
		},
	},
];
