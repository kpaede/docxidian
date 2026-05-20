import { Notice, TFile, setIcon } from 'obsidian';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useRef } from 'react';
import { DocxEditor, DocxEditorRef } from '@eigenpal/docx-editor-react';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import editorStyles from '@eigenpal/docx-editor-react/styles.css';

let stylesInjected = false;

export function ensureEditorStyles() {
	if (stylesInjected) {
		return;
	}

	const styleSheet = new CSSStyleSheet();
	styleSheet.replaceSync(editorStyles);
	document.adoptedStyleSheets = [...document.adoptedStyleSheets, styleSheet];
	stylesInjected = true;
}

const SaveButton = ({ onClick }: { onClick: () => void }) => {
	const ref = useRef<HTMLButtonElement>(null);
	useEffect(() => {
		if (ref.current) {
			ref.current.innerHTML = '';
			setIcon(ref.current, 'save');
		}
	}, []);

	return (
		<button
			ref={ref}
			type="button"
			className="clickable-icon docxidian-logo-save-button"
			onClick={onClick}
			aria-label="Save"
			title="Save"
			style={{ 
				background: 'transparent',
				border: 'none',
				boxShadow: 'none',
				padding: '4px 8px',
				cursor: 'pointer',
				display: 'flex',
				alignItems: 'center',
				justifyContent: 'center',
				color: 'inherit'
			}}
		/>
	);
};

export interface DocxReactViewProps {
	file: TFile | null;
	buffer: ArrayBuffer | null;
	error: string | null;
	isLoading: boolean;
	authorName: string;
	i18n: Translations | undefined;
	onDirtyChange: (isDirty: boolean) => void;
	onSave: (buffer: ArrayBuffer) => Promise<void>;
}

export interface DocxReactViewHandle {
	save: () => Promise<boolean>;
}

export const DocxReactView = forwardRef<DocxReactViewHandle, DocxReactViewProps>(function DocxReactView(
	{ file, buffer, error, isLoading, authorName, i18n, onDirtyChange, onSave },
	ref,
) {
	const editorRef = useRef<DocxEditorRef>(null);
	const dirtyTrackingEnabledRef = useRef(false);
	const isSavingRef = useRef(false);

	useEffect(() => {
		ensureEditorStyles();
	}, []);

	useEffect(() => {
		dirtyTrackingEnabledRef.current = false;
		const timeout = window.setTimeout(() => {
			dirtyTrackingEnabledRef.current = true;
		}, 500);

		return () => {
			window.clearTimeout(timeout);
			dirtyTrackingEnabledRef.current = false;
		};
	}, [file, buffer]);

	const persistDocument = useCallback(async (output: ArrayBuffer) => {
		if (!file) {
			return false;
		}

		if (isSavingRef.current) {
			return false;
		}
		isSavingRef.current = true;

		try {
			await onSave(output);
			new Notice(`Saved ${file.name}`);
			return true;
		} catch (saveError) {
			const message = saveError instanceof Error ? saveError.message : 'Unknown save error';
			new Notice(`Could not save ${file.name}: ${message}`);
			return false;
		} finally {
			setTimeout(() => {
				isSavingRef.current = false;
			}, 300);
		}
	}, [file, onSave]);

	const saveDocument = useCallback(async () => {
		if (!file) {
			new Notice('No docx file is open.');
			return false;
		}

		const output = await editorRef.current?.save({ selective: false });
		if (!output) {
			new Notice(`Could not save ${file.name}: the editor did not return a document.`);
			return false;
		}

		return persistDocument(output);
	}, [file, persistDocument]);

	useImperativeHandle(ref, () => ({
		save: saveDocument,
	}), [saveDocument]);

	if (isLoading) {
		return null;
	}

	if (error) {
		return <div>{error}</div>;
	}

	if (!file || !buffer) {
		return null;
	}

	return (
		<DocxEditor
			key={`${file.path}-${file.stat.mtime}`}
			ref={editorRef}
			documentBuffer={buffer}
			mode="editing"
			author={authorName}
			i18n={i18n}
			documentName={file.basename}
			documentNameEditable={false}
			renderLogo={() => (
				<SaveButton onClick={() => void saveDocument()} />
			)}
			onChange={() => {
				if (dirtyTrackingEnabledRef.current) {
					onDirtyChange(true);
				}
			}}
			onSave={(output) => {
				void persistDocument(output);
			}}
			onError={(docxError) => {
				new Notice(`Could not render ${file.name}: ${docxError.message}`);
			}}
		/>
	);
});
