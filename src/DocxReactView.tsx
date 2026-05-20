import { Notice, TFile, setIcon } from 'obsidian';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState, type ComponentProps } from 'react';
import { DocxEditor, DocxEditorRef } from '@eigenpal/docx-editor-react';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import editorStyles from '@eigenpal/docx-editor-react/styles.css';

let stylesInjected = false;
let editorInstanceCounter = 0;

interface DocxSectionProperties {
	pageHeight?: number;
	marginTop?: number;
	marginBottom?: number;
}

interface DocxDocumentWithSectionProperties {
	package?: {
		document?: {
			finalSectionProperties?: DocxSectionProperties;
			sections?: Array<{
				properties?: DocxSectionProperties;
			}>;
		};
	};
}

const DEFAULT_PAGE_HEIGHT_TWIPS = 15840;
const DEFAULT_MARGIN_TWIPS = 1440;

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
	showRuler: boolean;
	autosave: boolean;
	reserveReviewSidebar: boolean;
	onDirtyChange: (isDirty: boolean) => void;
	onSave: (buffer: ArrayBuffer) => Promise<void>;
	onDocumentNameChange: (name: string) => Promise<void>;
}

export interface DocxReactViewHandle {
	save: () => Promise<boolean>;
}

export const DocxReactView = forwardRef<DocxReactViewHandle, DocxReactViewProps>(function DocxReactView(
	{ file, buffer, error, isLoading, authorName, i18n, showRuler, autosave, reserveReviewSidebar, onDirtyChange, onSave, onDocumentNameChange },
	ref,
) {
	const editorRef = useRef<DocxEditorRef>(null);
	const editorClassNameRef = useRef(`docxidian-editor-${++editorInstanceCounter}`);
	const rulerSyncFrameRef = useRef<number | null>(null);
	const rulerSyncTimeoutRef = useRef<number | null>(null);
	const dirtyTrackingEnabledRef = useRef(false);
	const isSavingRef = useRef(false);
	const autosaveTimeoutRef = useRef<number | null>(null);
	const renameTimeoutRef = useRef<number | null>(null);
	const pendingSaveOptionsRef = useRef<{ silent?: boolean } | undefined>(undefined);
	const pendingSavePromiseRef = useRef<Promise<boolean> | null>(null);
	const [documentName, setDocumentName] = useState(file?.name ?? '');
	const pluginSidebarItems = useMemo<NonNullable<ComponentProps<typeof DocxEditor>['pluginSidebarItems']>>(() => {
		if (!reserveReviewSidebar) {
			return [];
		}

		return [{
			id: 'docxidian-review-sidebar-reservation',
			anchorPos: 1,
			estimatedHeight: 1,
			priority: Number.MAX_SAFE_INTEGER,
			render: () => null,
		}];
	}, [reserveReviewSidebar]);

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

	useEffect(() => {
		setDocumentName(file?.name ?? '');
	}, [file]);

	const clearAutosaveTimeout = useCallback(() => {
		if (autosaveTimeoutRef.current !== null) {
			window.clearTimeout(autosaveTimeoutRef.current);
			autosaveTimeoutRef.current = null;
		}
	}, []);

	const clearRenameTimeout = useCallback(() => {
		if (renameTimeoutRef.current !== null) {
			window.clearTimeout(renameTimeoutRef.current);
			renameTimeoutRef.current = null;
		}
	}, []);

	const syncVerticalRulerMarkers = useCallback((docxDocument: DocxDocumentWithSectionProperties | null | undefined) => {
		if (!showRuler || !docxDocument) {
			return;
		}

		const documentProperties = docxDocument.package?.document;
		const sectionProperties = {
			...documentProperties?.sections?.[0]?.properties,
			...documentProperties?.finalSectionProperties,
		};
		const pageHeight = sectionProperties.pageHeight ?? DEFAULT_PAGE_HEIGHT_TWIPS;
		const topMargin = sectionProperties.marginTop ?? DEFAULT_MARGIN_TWIPS;
		const bottomMargin = sectionProperties.marginBottom ?? DEFAULT_MARGIN_TWIPS;
		const ruler = document.querySelector<HTMLElement>(`.${editorClassNameRef.current} .docx-vertical-ruler`);

		if (!ruler || pageHeight <= 0) {
			return;
		}

		const pxPerTwip = ruler.getBoundingClientRect().height / pageHeight;
		const topMarker = ruler.querySelector<HTMLElement>('.docx-ruler-marker-topMargin');
		const bottomMarker = ruler.querySelector<HTMLElement>('.docx-ruler-marker-bottomMargin');

		if (topMarker) {
			topMarker.style.top = `${Math.round(topMargin * pxPerTwip - 5)}px`;
		}
		if (bottomMarker) {
			bottomMarker.style.top = `${Math.round((pageHeight - bottomMargin) * pxPerTwip - 5)}px`;
		}
	}, [showRuler]);

	const scheduleVerticalRulerMarkerSync = useCallback((document: DocxDocumentWithSectionProperties | null | undefined) => {
		if (rulerSyncFrameRef.current !== null) {
			window.cancelAnimationFrame(rulerSyncFrameRef.current);
		}
		if (rulerSyncTimeoutRef.current !== null) {
			window.clearTimeout(rulerSyncTimeoutRef.current);
		}

		rulerSyncFrameRef.current = window.requestAnimationFrame(() => {
			rulerSyncFrameRef.current = null;
			syncVerticalRulerMarkers(document);
			window.requestAnimationFrame(() => syncVerticalRulerMarkers(document));
		});
		rulerSyncTimeoutRef.current = window.setTimeout(() => {
			rulerSyncTimeoutRef.current = null;
			syncVerticalRulerMarkers(document);
		}, 50);
	}, [syncVerticalRulerMarkers]);

	useEffect(() => {
		if (showRuler) {
			scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
		}
	}, [showRuler, file, buffer, scheduleVerticalRulerMarkerSync]);

	useEffect(() => () => {
		clearAutosaveTimeout();
		clearRenameTimeout();
		if (rulerSyncFrameRef.current !== null) {
			window.cancelAnimationFrame(rulerSyncFrameRef.current);
			rulerSyncFrameRef.current = null;
		}
		if (rulerSyncTimeoutRef.current !== null) {
			window.clearTimeout(rulerSyncTimeoutRef.current);
			rulerSyncTimeoutRef.current = null;
		}
	}, [clearAutosaveTimeout, clearRenameTimeout]);

	const persistDocument = useCallback(async (output: ArrayBuffer, options?: { silent?: boolean }) => {
		if (!file) {
			return false;
		}

		if (isSavingRef.current) {
			return false;
		}
		isSavingRef.current = true;

		try {
			await onSave(output);
			onDirtyChange(false);
			if (!options?.silent) {
				new Notice(`Saved ${file.name}`);
			}
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
	}, [file, onDirtyChange, onSave]);

	const saveDocument = useCallback(async (options?: { silent?: boolean }) => {
		clearAutosaveTimeout();

		if (!file) {
			new Notice('No docx file is open.');
			return false;
		}

		pendingSaveOptionsRef.current = options;
		pendingSavePromiseRef.current = null;
		const output = await editorRef.current?.save({ selective: false });
		const pendingSavePromise = pendingSavePromiseRef.current;
		pendingSaveOptionsRef.current = undefined;
		pendingSavePromiseRef.current = null;

		if (!output) {
			new Notice(`Could not save ${file.name}: the editor did not return a document.`);
			return false;
		}

		if (pendingSavePromise) {
			return pendingSavePromise;
		}

		return persistDocument(output, options);
	}, [clearAutosaveTimeout, file, persistDocument]);

	useEffect(() => {
		if (!autosave) {
			clearAutosaveTimeout();
		}
	}, [autosave, clearAutosaveTimeout]);

	const scheduleAutosave = useCallback(() => {
		if (!autosave) {
			clearAutosaveTimeout();
			return;
		}

		clearAutosaveTimeout();
		autosaveTimeoutRef.current = window.setTimeout(() => {
			autosaveTimeoutRef.current = null;
			void saveDocument({ silent: true });
		}, 1500);
	}, [autosave, clearAutosaveTimeout, saveDocument]);

	const scheduleRename = useCallback((name: string) => {
		clearRenameTimeout();
		renameTimeoutRef.current = window.setTimeout(async () => {
			renameTimeoutRef.current = null;
			try {
				await onDocumentNameChange(name);
			} catch (renameError) {
				const message = renameError instanceof Error ? renameError.message : 'Unknown rename error';
				new Notice(`Could not rename ${file?.name ?? 'document'}: ${message}`);
				setDocumentName(file?.name ?? '');
			}
		}, 700);
	}, [clearRenameTimeout, file, onDocumentNameChange]);

	useImperativeHandle(ref, () => ({
		save: () => saveDocument(),
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
			className={editorClassNameRef.current}
			showRuler={showRuler}
			documentName={documentName}
			documentNameEditable
			pluginSidebarItems={pluginSidebarItems.length > 0 ? pluginSidebarItems : undefined}
			onDocumentNameChange={(name) => {
				setDocumentName(name);
				scheduleRename(name);
			}}
			renderLogo={() => (
				<SaveButton onClick={() => void saveDocument()} />
			)}
			onChange={() => {
				if (dirtyTrackingEnabledRef.current) {
					onDirtyChange(true);
					scheduleAutosave();
				}
				scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
			}}
			onSave={(output) => {
				const savePromise = persistDocument(output, pendingSaveOptionsRef.current);
				pendingSavePromiseRef.current = savePromise;
				void savePromise;
			}}
			onError={(docxError) => {
				new Notice(`Could not render ${file.name}: ${docxError.message}`);
			}}
		/>
	);
});
