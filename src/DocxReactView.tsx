import { Notice, TFile } from 'obsidian';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useRef } from 'react';
import { DocxEditor, DocxEditorRef } from '@eigenpal/docx-editor-react';
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

export interface DocxReactViewProps {
	file: TFile | null;
	buffer: ArrayBuffer | null;
	error: string | null;
	isLoading: boolean;
	authorName: string;
	onDirtyChange: (isDirty: boolean) => void;
	onSave: (buffer: ArrayBuffer) => Promise<void>;
}

export interface DocxReactViewHandle {
	save: () => Promise<boolean>;
}

export const DocxReactView = forwardRef<DocxReactViewHandle, DocxReactViewProps>(function DocxReactView(
	{ file, buffer, error, isLoading, authorName, onDirtyChange, onSave },
	ref,
) {
	const editorRef = useRef<DocxEditorRef>(null);
	const dirtyTrackingEnabledRef = useRef(false);

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

		try {
			await onSave(output);
			new Notice(`Saved ${file.name}`);
			return true;
		} catch (saveError) {
			const message = saveError instanceof Error ? saveError.message : 'Unknown save error';
			new Notice(`Could not save ${file.name}: ${message}`);
			return false;
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
			documentName={file.basename}
			documentNameEditable={false}
			renderTitleBarRight={() => (
				<button
					type="button"
					className="docxidian-title-save-button"
					onClick={() => {
						void saveDocument();
					}}
				>
					Speichern
				</button>
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
