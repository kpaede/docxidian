import { App, FileView, Modal, Notice, TFile, WorkspaceLeaf } from 'obsidian';
import { createRef } from 'react';
import { Root, createRoot } from 'react-dom/client';
import type { Translations } from '@eigenpal/docx-editor-i18n';
import { DocxReactView, DocxReactViewHandle } from './DocxReactView';
import { hasReviewMarkup } from './docxReviewMarkup';

export const VIEW_TYPE_DOCX = 'docxidian-docx-view';

type UnsavedDocxChoice = 'save' | 'discard';

class UnsavedDocxModal extends Modal {
	private resolveChoice: (choice: UnsavedDocxChoice) => void;
	private resolved = false;

	constructor(
		app: App,
		private fileName: string,
		resolveChoice: (choice: UnsavedDocxChoice) => void,
	) {
		super(app);
		this.resolveChoice = resolveChoice;
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Änderungen speichern?' });
		contentEl.createEl('p', { text: `${this.fileName} hat ungespeicherte Änderungen.` });

		const buttonRow = contentEl.createDiv({ cls: 'docxidian-unsaved-actions' });
		const discardButton = buttonRow.createEl('button', { text: 'Verwerfen' });
		const saveButton = buttonRow.createEl('button', { text: 'Speichern' });
		saveButton.addClass('mod-cta');

		discardButton.addEventListener('click', () => this.choose('discard'));
		saveButton.addEventListener('click', () => this.choose('save'));
	}

	onClose() {
		this.contentEl.empty();
		if (!this.resolved) {
			this.choose('discard');
		}
	}

	private choose(choice: UnsavedDocxChoice) {
		if (this.resolved) {
			return;
		}

		this.resolved = true;
		this.resolveChoice(choice);
		this.close();
	}
}

function shouldHandleEditorSaveClick(target: EventTarget | null) {
	if (!(target instanceof Element)) {
		return false;
	}

	let candidate: Element | null = target;
	while (candidate && candidate !== document.body) {
		if (candidate instanceof HTMLElement) {
			const text = candidate.textContent?.replace(/\s+/g, ' ').trim().toLowerCase() ?? '';
			if (/^save(?:\b|⌘|ctrl|\s)/.test(text) || /^speichern(?:\b|⌘|ctrl|\s)/.test(text)) {
				return true;
			}
		}

		candidate = candidate.parentElement;
	}

	return false;
}

export class DocxView extends FileView {
	private root: Root | null = null;
	private hostEl: HTMLDivElement | null = null;
	private reactViewRef = createRef<DocxReactViewHandle>();
	private buffer: ArrayBuffer | null = null;
	private error: string | null = null;
	private isLoading = false;
	private isDirty = false;
	private reserveReviewSidebar = false;
	private hostResizeObserver: ResizeObserver | null = null;
	private titleObserver: MutationObserver | null = null;
	private helpMenuObserver: MutationObserver | null = null;

	constructor(
		leaf: WorkspaceLeaf,
		private getAuthorName: () => string,
		private getEditorLocale: () => Translations | undefined,
		private getShowRuler: () => boolean,
		private getAutosave: () => boolean,
	) {
		super(leaf);
	}

	getViewType() {
		return VIEW_TYPE_DOCX;
	}

	getDisplayText() {
		return this.file?.basename ?? 'DOCX';
	}

	getIcon() {
		return 'file-text';
	}

	canAcceptExtension(extension: string) {
		return extension.toLowerCase() === 'docx';
	}

	async onOpen() {
		this.contentEl.empty();
		this.hostEl = this.contentEl.createDiv({ cls: 'docxidian-host' });
		this.prepareViewHost();
		this.registerHostMetrics();
		this.removeNativeButtonTitles();
		this.removeEditorHelpMenu();
		this.trackEditorHoverState();
		this.registerEditorSaveInterceptor();
		this.registerSaveShortcut();
		this.registerEditorDropdownScrollGuard();
		this.root = createRoot(this.hostEl);
		this.render();
	}

	async onClose() {
		await this.promptToSaveIfDirty();
		this.root?.unmount();
		this.root = null;
		this.hostResizeObserver?.disconnect();
		this.hostResizeObserver = null;
		this.titleObserver?.disconnect();
		this.titleObserver = null;
		this.helpMenuObserver?.disconnect();
		this.helpMenuObserver = null;
		document.body.classList.remove('docxidian-editor-hovering');
		this.hostEl = null;
		this.reactViewRef = createRef<DocxReactViewHandle>();
		this.buffer = null;
		this.error = null;
		this.isDirty = false;
		this.reserveReviewSidebar = false;
	}

	async onLoadFile(file: TFile) {
		await this.promptToSaveIfDirty();
		this.isLoading = true;
		this.error = null;
		this.buffer = null;
		this.isDirty = false;
		this.reserveReviewSidebar = false;
		this.render();

		try {
			this.buffer = await this.app.vault.readBinary(file);
			this.reserveReviewSidebar = await hasReviewMarkup(this.buffer);
		} catch (readError) {
			const message = readError instanceof Error ? readError.message : 'Unknown read error';
			this.error = `Could not load ${file.name}: ${message}`;
			new Notice(this.error);
		} finally {
			this.isLoading = false;
			this.render();
		}
	}

	async onUnloadFile(_file: TFile) {
		await this.promptToSaveIfDirty();
		this.buffer = null;
		this.error = null;
		this.isDirty = false;
		this.reserveReviewSidebar = false;
		this.render();
	}

	async onRename(file: TFile) {
		await super.onRename(file);
		this.render();
	}

	async saveCurrentDocument() {
		if (!this.file) {
			new Notice('No docx file is open.');
			return false;
		}

		if (this.isLoading) {
			new Notice(`Still loading ${this.file.name}.`);
			return false;
		}

		const saved = await this.reactViewRef.current?.save() ?? false;
		if (saved) {
			this.isDirty = false;
		}

		return saved;
	}

	refreshSettings() {
		this.render();
	}

	private async saveFile(buffer: ArrayBuffer) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}

		await this.app.vault.modifyBinary(file, buffer);
		this.isDirty = false;
	}

	private async renameFile(name: string) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}

		const normalizedName = this.normalizeDocxFileName(name);
		if (!normalizedName) {
			throw new Error('Document name cannot be empty.');
		}

		if (normalizedName === file.name) {
			return;
		}

		const folderPath = file.parent?.path;
		const newPath = folderPath && folderPath !== '/' ? `${folderPath}/${normalizedName}` : normalizedName;
		await this.app.fileManager.renameFile(file, newPath);
		new Notice(`Renamed to ${normalizedName}`);
	}

	private normalizeDocxFileName(name: string) {
		const trimmedName = name.trim().replace(/[\\/]/g, '-');
		if (!trimmedName || trimmedName === '.docx') {
			return null;
		}

		return trimmedName.toLowerCase().endsWith('.docx') ? trimmedName : `${trimmedName}.docx`;
	}

	private async promptToSaveIfDirty() {
		if (!this.isDirty || !this.file) {
			return;
		}

		const choice = await new Promise<UnsavedDocxChoice>((resolve) => {
			new UnsavedDocxModal(this.app, this.file?.name ?? 'Document', resolve).open();
		});

		if (choice === 'save') {
			await this.saveCurrentDocument();
		}
	}

	private prepareViewHost() {
		this.contentEl.setCssProps({
			padding: '0',
			overflow: 'hidden',
			height: '100%',
			display: 'flex',
			'flex-direction': 'column',
			'min-height': '0',
		});

		if (!this.hostEl) {
			return;
		}

		this.hostEl.setCssProps({
			'--docxidian-fixed-left-offset': '0px',
			'--docxidian-fixed-top-offset': '0px',
			background: '#f8fafc',
			display: 'flex',
			flex: '1 1 auto',
			'flex-direction': 'column',
			height: '100%',
			'min-height': '0',
			overflow: 'hidden',
			width: '100%',
		});
	}

	private registerHostMetrics() {
		const updateHostMetrics = () => {
			if (!this.hostEl) {
				return;
			}

			const fixedProbe = this.hostEl.createDiv({ cls: 'docxidian-fixed-probe' });
			fixedProbe.setCssProps({
				left: '0',
				'pointer-events': 'none',
				position: 'fixed',
				top: '0',
				visibility: 'hidden',
			});
			const fixedRect = fixedProbe.getBoundingClientRect();
			fixedProbe.remove();

			this.hostEl.setCssProps({
				'--docxidian-fixed-left-offset': `${Math.round(fixedRect.left)}px`,
				'--docxidian-fixed-top-offset': `${Math.round(fixedRect.top)}px`,
			});
		};

		updateHostMetrics();
		this.registerDomEvent(window, 'resize', updateHostMetrics);
		this.registerDomEvent(window, 'scroll', updateHostMetrics, true);
		this.hostResizeObserver = new ResizeObserver(updateHostMetrics);
		this.hostResizeObserver.observe(this.contentEl);
		this.register(() => {
			this.hostResizeObserver?.disconnect();
			this.hostResizeObserver = null;
		});
	}

	private removeNativeButtonTitles() {
		if (!this.hostEl) {
			return;
		}

		const removeTitles = () => {
			this.hostEl?.querySelectorAll('.ep-root button[title]').forEach((button) => {
				button.removeAttribute('title');
			});
		};

		removeTitles();
		this.titleObserver = new MutationObserver(removeTitles);
		this.titleObserver.observe(this.hostEl, {
			attributes: true,
			attributeFilter: ['title'],
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.titleObserver?.disconnect();
			this.titleObserver = null;
		});
	}

	private removeEditorHelpMenu() {
		if (!this.hostEl) {
			return;
		}

		const removeHelpMenu = () => {
			this.hostEl?.querySelectorAll('.ep-root [role="menubar"] > div').forEach((menuItem) => {
				const button = menuItem.querySelector(':scope > button');
				const label = button?.textContent?.replace(/\s+/g, ' ').trim().toLowerCase();
				if (label === 'help' || label === 'hilfe') {
					menuItem.remove();
				}
			});
		};

		removeHelpMenu();
		this.helpMenuObserver = new MutationObserver(removeHelpMenu);
		this.helpMenuObserver.observe(this.hostEl, {
			childList: true,
			subtree: true,
		});
		this.register(() => {
			this.helpMenuObserver?.disconnect();
			this.helpMenuObserver = null;
		});
	}

	private trackEditorHoverState() {
		if (!this.hostEl) {
			return;
		}

		const markEditorHovering = () => document.body.classList.add('docxidian-editor-hovering');
		const clearEditorHovering = () => document.body.classList.remove('docxidian-editor-hovering');

		this.registerDomEvent(this.hostEl, 'pointerenter', markEditorHovering);
		this.registerDomEvent(this.hostEl, 'pointerleave', clearEditorHovering);
		this.register(clearEditorHovering);
	}

	private registerEditorSaveInterceptor() {
		this.registerDomEvent(document, 'click', (evt) => {
			if (
				!this.hostEl
				|| this.app.workspace.getActiveViewOfType(DocxView) !== this
				|| (evt.target instanceof Element && !!evt.target.closest('.modal'))
				|| !shouldHandleEditorSaveClick(evt.target)
			) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.saveCurrentDocument();
		}, true);
	}

	private registerSaveShortcut() {
		this.registerDomEvent(document, 'keydown', (evt) => {
			if (
				!this.hostEl
				|| evt.key.toLowerCase() !== 's'
				|| (!evt.metaKey && !evt.ctrlKey)
				|| !(document.activeElement instanceof Node)
				|| !this.hostEl.contains(document.activeElement)
			) {
				return;
			}

			evt.preventDefault();
			evt.stopImmediatePropagation();
			void this.saveCurrentDocument();
		}, true);
	}

	private registerEditorDropdownScrollGuard() {
		const keepEditorListboxOpen = (evt: Event) => {
			if (!this.hostEl || this.app.workspace.getActiveViewOfType(DocxView) !== this || !(evt.target instanceof Element)) {
				return;
			}

			const listbox = evt.target.closest('[role="listbox"]');
			if (listbox && this.hostEl.contains(listbox)) {
				evt.stopImmediatePropagation();
				evt.stopPropagation();
			}
		};

		this.registerDomEvent(window, 'scroll', keepEditorListboxOpen, true);
	}

	private render() {
		this.root?.render(
			<DocxReactView
				ref={this.reactViewRef}
				file={this.file}
				buffer={this.buffer}
				error={this.error}
				isLoading={this.isLoading}
				authorName={this.getAuthorName()}
				i18n={this.getEditorLocale()}
				showRuler={this.getShowRuler()}
				autosave={this.getAutosave()}
				reserveReviewSidebar={this.reserveReviewSidebar}
				onDirtyChange={(isDirty) => {
					this.isDirty = isDirty;
				}}
				onSave={(buffer) => this.saveFile(buffer)}
				onDocumentNameChange={(name) => this.renameFile(name)}
			/>,
		);
	}
}
