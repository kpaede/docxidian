import editorStyles from '@eigenpal/docx-editor-react/styles.css';

let styleSheet: CSSStyleSheet | null = null;

export function ensureEditorStyles() {
	if (styleSheet) {
		return;
	}

	styleSheet = new CSSStyleSheet();
	styleSheet.replaceSync(editorStyles);
	document.adoptedStyleSheets = [...document.adoptedStyleSheets, styleSheet];
}

export function removeEditorStyles() {
	if (!styleSheet) {
		return;
	}

	document.adoptedStyleSheets = document.adoptedStyleSheets.filter(sheet => sheet !== styleSheet);
	styleSheet = null;
}
