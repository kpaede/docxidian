import { Notice, Platform, TFile, setIcon } from 'obsidian';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState, type ComponentProps } from 'react';
import { DocxEditor, type DocxEditorRef, type EditorMode } from '@eigenpal/docx-editor-react';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import editorStyles from '@eigenpal/docx-editor-react/styles.css';
import type { DocxidianLanguage } from './locales';

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
const MIN_TOUCH_ZOOM = 0.25;
const MAX_TOUCH_ZOOM = 4;
const TOUCH_ZOOM_SENSITIVITY = 0.55;
const TOUCH_ZOOM_MIN_DELTA = 0.006;

type FindReplaceMode = 'find' | 'replace';

interface FindMatch {
	from: number;
	to: number;
	text: string;
}

interface RefreshFindOptions {
	select?: boolean;
}

interface FindHighlightState {
	matches: FindMatch[];
	currentIndex: number;
}

interface PinchZoomState {
	source: 'touch' | 'gesture' | 'pointer';
	startDistance: number;
	lastDistance: number;
	startZoom: number;
	lastZoom: number;
}

interface PointerPoint {
	x: number;
	y: number;
}

type WebKitGestureEvent = Event & {
	clientX?: number;
	clientY?: number;
	scale?: number;
};

interface FindReplaceLabels {
	find: string;
	findAndReplace: string;
	findText: string;
	replaceWith: string;
	replace: string;
	replaceAll: string;
	matchCase: string;
	wholeWords: string;
	showReplace: string;
	close: string;
	previous: string;
	next: string;
	noMatches: string;
	resultCount: (current: number, total: number) => string;
}

const FIND_REPLACE_LABELS: Record<DocxidianLanguage, FindReplaceLabels> = {
	en: {
		find: 'Find',
		findAndReplace: 'Find and Replace',
		findText: 'Find text',
		replaceWith: 'Replace with',
		replace: 'Replace',
		replaceAll: 'Replace all',
		matchCase: 'Match case',
		wholeWords: 'Whole words',
		showReplace: 'Show replace',
		close: 'Close',
		previous: 'Previous match',
		next: 'Next match',
		noMatches: 'No matches',
		resultCount: (current, total) => `${current} of ${total}`,
	},
	de: {
		find: 'Suchen',
		findAndReplace: 'Suchen und Ersetzen',
		findText: 'Suchtext',
		replaceWith: 'Ersetzen durch',
		replace: 'Ersetzen',
		replaceAll: 'Alle ersetzen',
		matchCase: 'Groß/Klein',
		wholeWords: 'Ganze Wörter',
		showReplace: 'Ersetzen anzeigen',
		close: 'Schließen',
		previous: 'Vorheriger Treffer',
		next: 'Nächster Treffer',
		noMatches: 'Keine Treffer',
		resultCount: (current, total) => `${current} von ${total}`,
	},
	pl: {
		find: 'Znajdź',
		findAndReplace: 'Znajdź i zamień',
		findText: 'Szukany tekst',
		replaceWith: 'Zamień na',
		replace: 'Zamień',
		replaceAll: 'Zamień wszystko',
		matchCase: 'Uwzględnij wielkość liter',
		wholeWords: 'Całe wyrazy',
		showReplace: 'Pokaż zamianę',
		close: 'Zamknij',
		previous: 'Poprzednie trafienie',
		next: 'Następne trafienie',
		noMatches: 'Brak trafień',
		resultCount: (current, total) => `${current} z ${total}`,
	},
	'pt-BR': {
		find: 'Localizar',
		findAndReplace: 'Localizar e substituir',
		findText: 'Texto para localizar',
		replaceWith: 'Substituir por',
		replace: 'Substituir',
		replaceAll: 'Substituir tudo',
		matchCase: 'Diferenciar maiúsculas',
		wholeWords: 'Palavras inteiras',
		showReplace: 'Mostrar substituir',
		close: 'Fechar',
		previous: 'Resultado anterior',
		next: 'Próximo resultado',
		noMatches: 'Nenhum resultado',
		resultCount: (current, total) => `${current} de ${total}`,
	},
	tr: {
		find: 'Bul',
		findAndReplace: 'Bul ve değiştir',
		findText: 'Aranacak metin',
		replaceWith: 'Şununla değiştir',
		replace: 'Değiştir',
		replaceAll: 'Tümünü değiştir',
		matchCase: 'Büyük/küçük harf',
		wholeWords: 'Tam sözcükler',
		showReplace: 'Değiştirmeyi göster',
		close: 'Kapat',
		previous: 'Önceki eşleşme',
		next: 'Sonraki eşleşme',
		noMatches: 'Eşleşme yok',
		resultCount: (current, total) => `${current} / ${total}`,
	},
	he: {
		find: 'חיפוש',
		findAndReplace: 'חיפוש והחלפה',
		findText: 'טקסט לחיפוש',
		replaceWith: 'החלפה ב',
		replace: 'החלף',
		replaceAll: 'החלף הכל',
		matchCase: 'התאם אותיות גדולות/קטנות',
		wholeWords: 'מילים שלמות',
		showReplace: 'הצג החלפה',
		close: 'סגור',
		previous: 'התוצאה הקודמת',
		next: 'התוצאה הבאה',
		noMatches: 'אין תוצאות',
		resultCount: (current, total) => `${current} מתוך ${total}`,
	},
	'zh-CN': {
		find: '查找',
		findAndReplace: '查找和替换',
		findText: '查找文本',
		replaceWith: '替换为',
		replace: '替换',
		replaceAll: '全部替换',
		matchCase: '区分大小写',
		wholeWords: '全字匹配',
		showReplace: '显示替换',
		close: '关闭',
		previous: '上一个匹配项',
		next: '下一个匹配项',
		noMatches: '无匹配项',
		resultCount: (current, total) => `${current} / ${total}`,
	},
};

const findHighlightPluginKey = new PluginKey<FindHighlightState>('docxidian-find-highlight');

function clampZoom(zoom: number) {
	return Math.max(MIN_TOUCH_ZOOM, Math.min(MAX_TOUCH_ZOOM, Math.round(zoom * 100) / 100));
}

function scaleTouchZoom(startZoom: number, rawScale: number) {
	if (!Number.isFinite(rawScale) || rawScale <= 0) {
		return startZoom;
	}

	return clampZoom(startZoom * Math.pow(rawScale, TOUCH_ZOOM_SENSITIVITY));
}

function getTouchDistance(first: Touch, second: Touch) {
	return Math.hypot(first.clientX - second.clientX, first.clientY - second.clientY);
}

function getTouchCenter(first: Touch, second: Touch) {
	return {
		x: (first.clientX + second.clientX) / 2,
		y: (first.clientY + second.clientY) / 2,
	};
}

function getPointDistance(first: PointerPoint, second: PointerPoint) {
	return Math.hypot(first.x - second.x, first.y - second.y);
}

function getPointCenter(first: PointerPoint, second: PointerPoint) {
	return {
		x: (first.x + second.x) / 2,
		y: (first.y + second.y) / 2,
	};
}

function getScrollableEditorElement(root: HTMLElement) {
	const pages = root.querySelector<HTMLElement>('.paged-editor__pages');
	let candidate: HTMLElement | null = pages?.parentElement ?? root;

	while (candidate && candidate !== root) {
		const style = window.getComputedStyle(candidate);
		const canScroll = /(auto|scroll)/.test(`${style.overflow} ${style.overflowX} ${style.overflowY}`);
		if (canScroll && (candidate.scrollHeight > candidate.clientHeight || candidate.scrollWidth > candidate.clientWidth)) {
			return candidate;
		}
		candidate = candidate.parentElement;
	}

	return root;
}

function shouldEnableTouchPinchZoom() {
	if (typeof window === 'undefined' || typeof navigator === 'undefined') {
		return false;
	}

	return Platform.isMobile || Platform.isMobileApp || (navigator.maxTouchPoints >= 2 && window.matchMedia('(hover: none)').matches);
}

function getEditorModeFromButton(button: HTMLButtonElement): EditorMode | null {
	const label = button.textContent?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '';

	if (label.startsWith('bearbeiten') || label.startsWith('edit')) {
		return 'editing';
	}
	if (label.startsWith('vorschlagen') || label.startsWith('suggest')) {
		return 'suggesting';
	}
	if (label.startsWith('anzeigen') || label.startsWith('view')) {
		return 'viewing';
	}

	return null;
}

function escapeRegExp(value: string) {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function createFindPattern(searchText: string, matchCase: boolean, wholeWord: boolean) {
	if (!searchText.trim()) {
		return null;
	}

	const source = wholeWord ? `\\b${escapeRegExp(searchText)}\\b` : escapeRegExp(searchText);
	return new RegExp(source, matchCase ? 'g' : 'gi');
}

function findMatchesInView(editor: DocxEditorRef | null, searchText: string, matchCase: boolean, wholeWord: boolean) {
	const view = editor?.getEditorRef()?.getView();
	const pattern = createFindPattern(searchText, matchCase, wholeWord);
	const matches: FindMatch[] = [];

	if (!view || !pattern) {
		return matches;
	}

	view.state.doc.descendants((node, pos) => {
		if (!node.isTextblock) {
			return true;
		}

		const text = node.textContent;
		for (const match of text.matchAll(pattern)) {
			const index = match.index ?? -1;
			if (index < 0) {
				continue;
			}

			matches.push({
				from: pos + 1 + index,
				to: pos + 1 + index + match[0].length,
				text: match[0],
			});
		}

		return false;
	});

	return matches;
}

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

const IconButton = ({ icon, label, onClick }: { icon: string; label: string; onClick: () => void }) => {
	const ref = useRef<HTMLButtonElement>(null);
	useEffect(() => {
		if (ref.current) {
			ref.current.innerHTML = '';
			setIcon(ref.current, icon);
		}
	}, [icon]);

	return (
		<button
			ref={ref}
			type="button"
			aria-label={label}
			title={label}
			onClick={onClick}
			style={{
				alignItems: 'center',
				background: 'transparent',
				border: 'none',
				borderRadius: '4px',
				boxShadow: 'none',
				color: 'inherit',
				cursor: 'pointer',
				display: 'inline-flex',
				height: '28px',
				justifyContent: 'center',
				padding: '4px 6px',
				width: '30px',
			}}
		/>
	);
};

interface FindReplaceDialogProps {
	isOpen: boolean;
	labels: FindReplaceLabels;
	mode: FindReplaceMode;
	searchText: string;
	replaceText: string;
	matchCase: boolean;
	wholeWord: boolean;
	matchCount: number;
	currentIndex: number;
	onSearchTextChange: (value: string) => void;
	onReplaceTextChange: (value: string) => void;
	onMatchCaseChange: (value: boolean) => void;
	onWholeWordChange: (value: boolean) => void;
	onModeChange: (mode: FindReplaceMode) => void;
	onNext: () => void;
	onPrevious: () => void;
	onReplace: () => void;
	onReplaceAll: () => void;
	onClose: () => void;
}

const FindReplaceDialog = ({
	isOpen,
	labels,
	mode,
	searchText,
	replaceText,
	matchCase,
	wholeWord,
	matchCount,
	currentIndex,
	onSearchTextChange,
	onReplaceTextChange,
	onMatchCaseChange,
	onWholeWordChange,
	onModeChange,
	onNext,
	onPrevious,
	onReplace,
	onReplaceAll,
	onClose,
}: FindReplaceDialogProps) => {
	if (!isOpen) {
		return null;
	}

	const resultText = searchText.trim()
		? (matchCount > 0 ? labels.resultCount(currentIndex + 1, matchCount) : labels.noMatches)
		: '';

	return (
		<div
			className="docxidian-find-dialog"
			style={{
				position: 'fixed',
				right: '24px',
				top: '92px',
				zIndex: 100050,
				width: '360px',
				background: 'white',
				border: '1px solid var(--background-modifier-border, #d1d5db)',
				borderRadius: '8px',
				boxShadow: '0 10px 30px rgba(0, 0, 0, 0.18)',
				padding: '12px',
				color: 'var(--text-normal, #202124)',
			}}
		>
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' }}>
				<strong>{mode === 'replace' ? labels.findAndReplace : labels.find}</strong>
				<button type="button" aria-label={labels.close} onClick={onClose} style={{ border: 'none', background: 'transparent', cursor: 'pointer' }}>×</button>
			</div>
			<div style={{ display: 'flex', gap: '6px', marginBottom: '8px' }}>
				<input
					value={searchText}
					onChange={(evt) => onSearchTextChange(evt.currentTarget.value)}
					placeholder={labels.findText}
					autoFocus
					style={{ flex: 1, height: '30px' }}
					onKeyDown={(evt) => {
						if (evt.key === 'Enter') {
							evt.preventDefault();
							evt.shiftKey ? onPrevious() : onNext();
						}
					}}
				/>
				<button type="button" aria-label={labels.previous} title={labels.previous} onClick={onPrevious} disabled={matchCount === 0}>↑</button>
				<button type="button" aria-label={labels.next} title={labels.next} onClick={onNext} disabled={matchCount === 0}>↓</button>
			</div>
			{mode === 'replace' && (
				<div style={{ display: 'flex', gap: '6px', marginBottom: '8px' }}>
					<input
						value={replaceText}
						onChange={(evt) => onReplaceTextChange(evt.currentTarget.value)}
						placeholder={labels.replaceWith}
						style={{ flex: 1, height: '30px' }}
					/>
					<button type="button" onClick={onReplace} disabled={matchCount === 0}>{labels.replace}</button>
					<button type="button" onClick={onReplaceAll} disabled={matchCount === 0}>{labels.replaceAll}</button>
				</div>
			)}
			<div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' }}>
				<div style={{ display: 'flex', gap: '10px', fontSize: '12px' }}>
					<label><input type="checkbox" checked={matchCase} onChange={(evt) => onMatchCaseChange(evt.currentTarget.checked)} /> {labels.matchCase}</label>
					<label><input type="checkbox" checked={wholeWord} onChange={(evt) => onWholeWordChange(evt.currentTarget.checked)} /> {labels.wholeWords}</label>
				</div>
				<div style={{ fontSize: '12px', color: 'var(--text-muted, #6b7280)', whiteSpace: 'nowrap' }}>{resultText}</div>
			</div>
			{mode === 'find' && (
				<button type="button" onClick={() => onModeChange('replace')} style={{ marginTop: '10px' }}>{labels.showReplace}</button>
			)}
		</div>
	);
};

export interface DocxReactViewProps {
	file: TFile | null;
	buffer: ArrayBuffer | null;
	error: string | null;
	isLoading: boolean;
	authorName: string;
	editorLanguage: DocxidianLanguage;
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
	openFind: () => void;
	openFindReplace: () => void;
}

export const DocxReactView = forwardRef<DocxReactViewHandle, DocxReactViewProps>(function DocxReactView(
	{ file, buffer, error, isLoading, authorName, editorLanguage, i18n, showRuler, autosave, reserveReviewSidebar, onDirtyChange, onSave, onDocumentNameChange },
	ref,
) {
	const editorRef = useRef<DocxEditorRef>(null);
	const editorClassNameRef = useRef(`docxidian-editor-${++editorInstanceCounter}`);
	const rulerSyncFrameRef = useRef<number | null>(null);
	const rulerSyncTimeoutRef = useRef<number | null>(null);
	const pinchZoomStateRef = useRef<PinchZoomState | null>(null);
	const pinchZoomScrollFrameRef = useRef<number | null>(null);
	const activeTouchPointersRef = useRef<Map<number, PointerPoint>>(new Map());
	const dirtyTrackingEnabledRef = useRef(false);
	const isSavingRef = useRef(false);
	const autosaveTimeoutRef = useRef<number | null>(null);
	const renameTimeoutRef = useRef<number | null>(null);
	const pendingSaveOptionsRef = useRef<{ silent?: boolean } | undefined>(undefined);
	const pendingSavePromiseRef = useRef<Promise<boolean> | null>(null);
	const [documentName, setDocumentName] = useState(file?.name ?? '');
	const [editorMode, setEditorMode] = useState<EditorMode>('editing');
	const [findDialogMode, setFindDialogMode] = useState<FindReplaceMode | null>(null);
	const [findSearchText, setFindSearchText] = useState('');
	const [findReplaceText, setFindReplaceText] = useState('');
	const [findMatchCase, setFindMatchCase] = useState(false);
	const [findWholeWord, setFindWholeWord] = useState(false);
	const [findMatches, setFindMatches] = useState<FindMatch[]>([]);
	const [currentFindIndex, setCurrentFindIndex] = useState(0);
	const filePath = file?.path ?? null;
	const findReplaceLabels = FIND_REPLACE_LABELS[editorLanguage] ?? FIND_REPLACE_LABELS.en;
	const findHighlightPlugin = useMemo(() => new Plugin<FindHighlightState>({
		key: findHighlightPluginKey,
		state: {
			init: () => ({ matches: [], currentIndex: 0 }),
			apply: (transaction, previous) => transaction.getMeta(findHighlightPluginKey) ?? previous,
		},
		props: {
			decorations: (state) => {
				const pluginState = findHighlightPluginKey.getState(state);
				if (!pluginState || pluginState.matches.length === 0) {
					return DecorationSet.empty;
				}

				return DecorationSet.create(
					state.doc,
					pluginState.matches.map((match, index) => Decoration.inline(
						match.from,
						match.to,
						{ class: index === pluginState.currentIndex ? 'docxidian-find-current' : 'docxidian-find-match' },
					)),
				);
			},
		},
	}), []);
	const externalPlugins = useMemo(() => [findHighlightPlugin], [findHighlightPlugin]);
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
		setEditorMode('editing');
		setFindDialogMode(null);
		setFindSearchText('');
		setFindReplaceText('');
		setFindMatches([]);
		setCurrentFindIndex(0);
	}, [filePath]);

	const setMode = useCallback((mode: EditorMode) => {
		setEditorMode(mode);
	}, []);

	const publishFindHighlights = useCallback((matches: FindMatch[], currentIndex: number) => {
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view) {
			return;
		}

		view.dispatch(view.state.tr.setMeta(findHighlightPluginKey, { matches, currentIndex }));
	}, []);

	const selectFindMatch = useCallback((matches: FindMatch[], index: number) => {
		const view = editorRef.current?.getEditorRef()?.getView();
		const match = matches[index];
		if (!view || !match) {
			return;
		}

		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, match.from, match.to)).scrollIntoView());
		editorRef.current?.scrollToPosition(match.from);
	}, []);

	const refreshFindMatches = useCallback((searchText: string, matchCase = findMatchCase, wholeWord = findWholeWord, preferredIndex = 0, options: RefreshFindOptions = {}) => {
		const matches = findMatchesInView(editorRef.current, searchText, matchCase, wholeWord);
		const nextIndex = matches.length > 0 ? Math.max(0, Math.min(preferredIndex, matches.length - 1)) : 0;

		setFindMatches(matches);
		setCurrentFindIndex(nextIndex);
		publishFindHighlights(matches, nextIndex);
		if (options.select && matches.length > 0) {
			selectFindMatch(matches, nextIndex);
		}

		return matches;
	}, [findMatchCase, findWholeWord, publishFindHighlights, selectFindMatch]);

	const openFindReplacePanel = useCallback((mode: FindReplaceMode) => {
		const selectedText = editorRef.current?.getSelectionInfo()?.selectedText?.trim();
		const nextSearchText = selectedText || findSearchText;

		setFindDialogMode(mode);
		if (selectedText) {
			setFindSearchText(selectedText);
		}
		refreshFindMatches(nextSearchText);
	}, [findSearchText, refreshFindMatches]);

	const openFindReplaceDialog = useCallback((mode: FindReplaceMode) => {
		openFindReplacePanel(mode);
	}, [openFindReplacePanel]);

	useEffect(() => {
		const handleFindShortcut = (evt: KeyboardEvent) => {
			const editorRoot = document.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			const isModifierPressed = evt.metaKey || evt.ctrlKey;
			if (
				!editorRoot
				|| !isModifierPressed
				|| evt.altKey
				|| evt.shiftKey
				|| !(evt.target instanceof Node)
				|| (!editorRoot.contains(evt.target) && !document.querySelector('.docxidian-find-dialog')?.contains(evt.target))
			) {
				return;
			}

			if (evt.key.toLowerCase() === 'f' || evt.key.toLowerCase() === 'h') {
				evt.preventDefault();
				evt.stopPropagation();
				openFindReplaceDialog(evt.key.toLowerCase() === 'f' ? 'find' : 'replace');
			}
		};

		document.addEventListener('keydown', handleFindShortcut, true);
		return () => document.removeEventListener('keydown', handleFindShortcut, true);
	}, [openFindReplaceDialog]);

	const moveFindMatch = useCallback((direction: 1 | -1) => {
		if (findMatches.length === 0) {
			return;
		}

		const nextIndex = (currentFindIndex + direction + findMatches.length) % findMatches.length;
		setCurrentFindIndex(nextIndex);
		publishFindHighlights(findMatches, nextIndex);
		selectFindMatch(findMatches, nextIndex);
	}, [currentFindIndex, findMatches, publishFindHighlights, selectFindMatch]);

	const replaceCurrentMatch = useCallback(() => {
		const view = editorRef.current?.getEditorRef()?.getView();
		const match = findMatches[currentFindIndex];
		if (!view || !match) {
			return;
		}

		const textNode = findReplaceText ? view.state.schema.text(findReplaceText) : null;
		view.dispatch(view.state.tr.replaceWith(match.from, match.to, textNode ? [textNode] : []).scrollIntoView());
		refreshFindMatches(findSearchText, findMatchCase, findWholeWord, currentFindIndex);
	}, [currentFindIndex, findMatchCase, findMatches, findReplaceText, findSearchText, findWholeWord, refreshFindMatches]);

	const replaceAllMatches = useCallback(() => {
		const view = editorRef.current?.getEditorRef()?.getView();
		if (!view || findMatches.length === 0) {
			return;
		}

		let transaction = view.state.tr;
		for (const match of [...findMatches].sort((a, b) => b.from - a.from)) {
			const textNode = findReplaceText ? view.state.schema.text(findReplaceText) : null;
			transaction = transaction.replaceWith(match.from, match.to, textNode ? [textNode] : []);
		}
		view.dispatch(transaction.scrollIntoView());
		refreshFindMatches(findSearchText, findMatchCase, findWholeWord, 0);
	}, [findMatchCase, findMatches, findReplaceText, findSearchText, findWholeWord, refreshFindMatches]);

	const normalizeEditorModeDropdown = useCallback(() => {
		const editorRoot = document.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		const modeMenus = Array.from(document.querySelectorAll<HTMLElement>('div[style*="position: fixed"]'))
			.map((menu) => ({
				menu,
				buttons: Array.from(menu.querySelectorAll<HTMLButtonElement>(':scope > button'))
					.filter((button) => getEditorModeFromButton(button) !== null && button.querySelector(':scope span span')),
			}))
			.filter(({ buttons }) => {
				const modes = new Set(buttons.map((button) => getEditorModeFromButton(button)));
				return buttons.length === 3 && modes.has('editing') && modes.has('suggesting') && modes.has('viewing');
			});

		modeMenus.forEach(({ menu, buttons }) => {
			menu.dataset.docxidianModeMenu = 'true';
			menu.style.minWidth = '260px';
			menu.style.padding = '4px 0';
			menu.style.overflow = 'hidden';

			buttons.forEach((button) => {
				const mode = getEditorModeFromButton(button);
				if (mode) {
					button.dataset.docxidianModeMenuItem = mode;
				}

				button.style.alignItems = 'center';
				button.style.boxShadow = 'none';
				button.style.columnGap = '10px';
				button.style.display = 'grid';
				button.style.gridTemplateColumns = '24px minmax(0, 1fr) 20px';
				button.style.justifyContent = 'start';
				button.style.justifyItems = 'start';
				button.style.lineHeight = 'normal';
				button.style.minHeight = '48px';
				button.style.padding = '8px 12px';
				button.style.width = '100%';

				const icon = button.querySelector<HTMLElement>(':scope > svg:first-child');
				if (icon) {
					icon.style.display = 'inline-flex';
					icon.style.flex = '0 0 20px';
					icon.style.gridColumn = '1';
					icon.style.height = '20px';
					icon.style.justifySelf = 'start';
					icon.style.marginLeft = '0';
					icon.style.width = '20px';
				}

				const labelColumn = button.querySelector<HTMLElement>(':scope > span');
				if (labelColumn) {
					labelColumn.style.flex = '1 1 auto';
					labelColumn.style.gridColumn = '2';
					labelColumn.style.justifySelf = 'stretch';
					labelColumn.style.marginLeft = '0';
					labelColumn.style.minWidth = '0';
				}

				const checkIcon = button.querySelector<HTMLElement>(':scope > svg:last-child:not(:first-child)');
				if (checkIcon) {
					checkIcon.style.gridColumn = '3';
					checkIcon.style.justifySelf = 'end';
					checkIcon.style.marginLeft = '0';
				}
			});
		});
	}, []);

	useEffect(() => {
		const editorRoot = document.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		normalizeEditorModeDropdown();
		const observer = new MutationObserver(normalizeEditorModeDropdown);
		observer.observe(document.body, {
			childList: true,
			subtree: true,
		});

		return () => observer.disconnect();
	}, [buffer, filePath, isLoading, normalizeEditorModeDropdown]);

	useEffect(() => {
		const handleModePointerDown = (evt: PointerEvent) => {
			const editorRoot = document.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
			if (!editorRoot || !(evt.target instanceof Element)) {
				return;
			}

			const button = evt.target.closest('button');
			if (!(button instanceof HTMLButtonElement)) {
				return;
			}

			const isEditorModeButton = editorRoot.contains(button) || button.closest('[data-docxidian-mode-menu]');
			if (!isEditorModeButton) {
				return;
			}
			const mode = getEditorModeFromButton(button);
			if (mode) {
				window.setTimeout(() => setMode(mode), 0);
			}
		};

		document.addEventListener('pointerdown', handleModePointerDown, true);
		return () => document.removeEventListener('pointerdown', handleModePointerDown, true);
	}, [setMode]);

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

	useEffect(() => {
		if (!shouldEnableTouchPinchZoom()) {
			return;
		}

		const editorRoot = document.querySelector<HTMLElement>(`.${editorClassNameRef.current}`);
		if (!editorRoot) {
			return;
		}

		const hostRoot = editorRoot.closest<HTMLElement>('.docxidian-host') ?? editorRoot;
		const previousTouchAction = editorRoot.style.touchAction;
		const previousHostTouchAction = hostRoot.style.touchAction;
		editorRoot.style.touchAction = 'pan-x pan-y';
		hostRoot.style.touchAction = 'pan-x pan-y';

		const isEditorTarget = (target: EventTarget | null) => target instanceof Node && hostRoot.contains(target);

		const shouldIgnoreGestureSource = (source: PinchZoomState['source']) => {
			const activeSource = pinchZoomStateRef.current?.source;
			return activeSource !== undefined && activeSource !== source;
		};

		const zoomAroundViewportPoint = (nextZoom: number, viewportPoint: PointerPoint, source: PinchZoomState['source']) => {
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== source || Math.abs(nextZoom - pinchState.lastZoom) < TOUCH_ZOOM_MIN_DELTA) {
				return false;
			}

			const scrollContainer = getScrollableEditorElement(editorRoot);
			const rect = scrollContainer.getBoundingClientRect();
			const localX = viewportPoint.x - rect.left;
			const localY = viewportPoint.y - rect.top;
			const documentX = (scrollContainer.scrollLeft + localX) / pinchState.lastZoom;
			const documentY = (scrollContainer.scrollTop + localY) / pinchState.lastZoom;

			editorRef.current?.setZoom(nextZoom);
			pinchState.lastZoom = nextZoom;

			if (pinchZoomScrollFrameRef.current !== null) {
				window.cancelAnimationFrame(pinchZoomScrollFrameRef.current);
			}
			pinchZoomScrollFrameRef.current = window.requestAnimationFrame(() => {
				pinchZoomScrollFrameRef.current = null;
				scrollContainer.scrollLeft = Math.max(0, documentX * nextZoom - localX);
				scrollContainer.scrollTop = Math.max(0, documentY * nextZoom - localY);
				scheduleVerticalRulerMarkerSync(editorRef.current?.getDocument());
			});
			return true;
		};

		const handleTouchStart = (evt: TouchEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('touch')) {
				return;
			}
			if (evt.touches.length !== 2) {
				if (pinchZoomStateRef.current?.source === 'touch') {
					pinchZoomStateRef.current = null;
				}
				return;
			}

			const first = evt.touches.item(0);
			const second = evt.touches.item(1);
			if (!first || !second) {
				return;
			}
			const startZoom = editorRef.current?.getZoom() ?? 1;
			const startDistance = getTouchDistance(first, second);
			if (startDistance <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			pinchZoomStateRef.current = {
				source: 'touch',
				startDistance,
				lastDistance: startDistance,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handleTouchMove = (evt: TouchEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== 'touch' || evt.touches.length !== 2) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();

			const first = evt.touches.item(0);
			const second = evt.touches.item(1);
			if (!first || !second) {
				return;
			}
			const distance = getTouchDistance(first, second);
			if (distance <= 0) {
				return;
			}

			const center = getTouchCenter(first, second);
			const didZoom = zoomAroundViewportPoint(scaleTouchZoom(pinchState.lastZoom, distance / pinchState.lastDistance), center, 'touch');
			if (didZoom) {
				pinchState.lastDistance = distance;
			}
		};

		const handleTouchEnd = (evt: TouchEvent) => {
			if (evt.touches.length < 2 && pinchZoomStateRef.current?.source === 'touch') {
				pinchZoomStateRef.current = null;
			}
		};

		const handleGestureEnd = () => {
			if (pinchZoomStateRef.current?.source === 'gesture') {
				pinchZoomStateRef.current = null;
			}
		};

		const handleGestureStart = (evt: WebKitGestureEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('gesture')) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const startZoom = editorRef.current?.getZoom() ?? 1;
			pinchZoomStateRef.current = {
				source: 'gesture',
				startDistance: 1,
				lastDistance: 1,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handleGestureChange = (evt: WebKitGestureEvent) => {
			if (!isEditorTarget(evt.target)) {
				return;
			}

			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || typeof evt.scale !== 'number' || evt.scale <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();

			const scrollContainer = getScrollableEditorElement(editorRoot);
			const rect = scrollContainer.getBoundingClientRect();
			zoomAroundViewportPoint(scaleTouchZoom(pinchState.startZoom, evt.scale), {
				x: evt.clientX ?? rect.left + rect.width / 2,
				y: evt.clientY ?? rect.top + rect.height / 2,
			}, 'gesture');
		};

		const handlePointerDown = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch' || !isEditorTarget(evt.target)) {
				return;
			}
			if (shouldIgnoreGestureSource('pointer')) {
				return;
			}

			activeTouchPointersRef.current.set(evt.pointerId, { x: evt.clientX, y: evt.clientY });
			if (activeTouchPointersRef.current.size !== 2) {
				return;
			}

			const [first, second] = Array.from(activeTouchPointersRef.current.values());
			if (!first || !second) {
				return;
			}
			const startDistance = getPointDistance(first, second);
			if (startDistance <= 0) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const startZoom = editorRef.current?.getZoom() ?? 1;
			pinchZoomStateRef.current = {
				source: 'pointer',
				startDistance,
				lastDistance: startDistance,
				startZoom,
				lastZoom: startZoom,
			};
		};

		const handlePointerMove = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch' || !activeTouchPointersRef.current.has(evt.pointerId)) {
				return;
			}

			activeTouchPointersRef.current.set(evt.pointerId, { x: evt.clientX, y: evt.clientY });
			const pinchState = pinchZoomStateRef.current;
			if (!pinchState || pinchState.source !== 'pointer' || activeTouchPointersRef.current.size !== 2) {
				return;
			}

			evt.preventDefault();
			evt.stopPropagation();
			const [first, second] = Array.from(activeTouchPointersRef.current.values());
			if (!first || !second) {
				return;
			}
			const distance = getPointDistance(first, second);
			if (distance <= 0) {
				return;
			}

			const didZoom = zoomAroundViewportPoint(
				scaleTouchZoom(pinchState.lastZoom, distance / pinchState.lastDistance),
				getPointCenter(first, second),
				'pointer',
			);
			if (didZoom) {
				pinchState.lastDistance = distance;
			}
		};

		const handlePointerEnd = (evt: PointerEvent) => {
			if (evt.pointerType !== 'touch') {
				return;
			}

			activeTouchPointersRef.current.delete(evt.pointerId);
			if (activeTouchPointersRef.current.size < 2 && pinchZoomStateRef.current?.source === 'pointer') {
				pinchZoomStateRef.current = null;
			}
		};

		document.addEventListener('touchstart', handleTouchStart, { passive: false, capture: true });
		document.addEventListener('touchmove', handleTouchMove, { passive: false, capture: true });
		document.addEventListener('touchend', handleTouchEnd, { passive: true, capture: true });
		document.addEventListener('touchcancel', handleTouchEnd, { passive: true, capture: true });
		document.addEventListener('gesturestart', handleGestureStart, { passive: false, capture: true });
		document.addEventListener('gesturechange', handleGestureChange, { passive: false, capture: true });
		document.addEventListener('gestureend', handleGestureEnd, { passive: true, capture: true });
		document.addEventListener('pointerdown', handlePointerDown, { passive: false, capture: true });
		document.addEventListener('pointermove', handlePointerMove, { passive: false, capture: true });
		document.addEventListener('pointerup', handlePointerEnd, { passive: true, capture: true });
		document.addEventListener('pointercancel', handlePointerEnd, { passive: true, capture: true });

		return () => {
			editorRoot.style.touchAction = previousTouchAction;
			hostRoot.style.touchAction = previousHostTouchAction;
			document.removeEventListener('touchstart', handleTouchStart, true);
			document.removeEventListener('touchmove', handleTouchMove, true);
			document.removeEventListener('touchend', handleTouchEnd, true);
			document.removeEventListener('touchcancel', handleTouchEnd, true);
			document.removeEventListener('gesturestart', handleGestureStart, true);
			document.removeEventListener('gesturechange', handleGestureChange, true);
			document.removeEventListener('gestureend', handleGestureEnd, true);
			document.removeEventListener('pointerdown', handlePointerDown, true);
			document.removeEventListener('pointermove', handlePointerMove, true);
			document.removeEventListener('pointerup', handlePointerEnd, true);
			document.removeEventListener('pointercancel', handlePointerEnd, true);
			pinchZoomStateRef.current = null;
			activeTouchPointersRef.current.clear();
			if (pinchZoomScrollFrameRef.current !== null) {
				window.cancelAnimationFrame(pinchZoomScrollFrameRef.current);
				pinchZoomScrollFrameRef.current = null;
			}
		};
	}, [buffer, filePath, isLoading, scheduleVerticalRulerMarkerSync]);

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
		openFind: () => openFindReplaceDialog('find'),
		openFindReplace: () => openFindReplaceDialog('replace'),
	}), [openFindReplaceDialog, saveDocument]);

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
		<>
			<DocxEditor
				key={`${file.path}-${file.stat.mtime}`}
				ref={editorRef}
				documentBuffer={buffer}
				mode={editorMode}
				onModeChange={setMode}
				author={authorName}
				i18n={i18n}
				className={editorClassNameRef.current}
				showRuler={showRuler}
				disableFindReplaceShortcuts
				externalPlugins={externalPlugins}
				documentName={documentName}
				documentNameEditable
				pluginSidebarItems={pluginSidebarItems.length > 0 ? pluginSidebarItems : undefined}
				toolbarExtra={(
					<>
						<IconButton icon="search" label={findReplaceLabels.find} onClick={() => openFindReplaceDialog('find')} />
						<IconButton icon="replace" label={findReplaceLabels.findAndReplace} onClick={() => openFindReplaceDialog('replace')} />
					</>
				)}
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
			<FindReplaceDialog
				isOpen={findDialogMode !== null}
				labels={findReplaceLabels}
				mode={findDialogMode ?? 'find'}
				searchText={findSearchText}
				replaceText={findReplaceText}
				matchCase={findMatchCase}
				wholeWord={findWholeWord}
				matchCount={findMatches.length}
				currentIndex={currentFindIndex}
				onSearchTextChange={(value) => {
					setFindSearchText(value);
					refreshFindMatches(value, findMatchCase, findWholeWord, 0);
				}}
				onReplaceTextChange={setFindReplaceText}
				onMatchCaseChange={(value) => {
					setFindMatchCase(value);
					refreshFindMatches(findSearchText, value, findWholeWord, currentFindIndex);
				}}
				onWholeWordChange={(value) => {
					setFindWholeWord(value);
					refreshFindMatches(findSearchText, findMatchCase, value, currentFindIndex);
				}}
				onModeChange={setFindDialogMode}
				onNext={() => moveFindMatch(1)}
				onPrevious={() => moveFindMatch(-1)}
				onReplace={replaceCurrentMatch}
				onReplaceAll={replaceAllMatches}
				onClose={() => {
					setFindDialogMode(null);
					setFindMatches([]);
					publishFindHighlights([], 0);
				}}
			/>
		</>
	);
});
