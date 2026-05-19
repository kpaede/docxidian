import { FileView, Notice, TFile, WorkspaceLeaf } from 'obsidian';
import { createRef } from 'react';
import { Root, createRoot } from 'react-dom/client';
import { DocxReactView, DocxReactViewHandle } from './DocxReactView';

export const VIEW_TYPE_DOCX = 'docxidian-docx-view';

export class DocxView extends FileView {
	private root: Root | null = null;
	private hostEl: HTMLDivElement | null = null;
	private reactViewRef = createRef<DocxReactViewHandle>();
	private buffer: ArrayBuffer | null = null;
	private error: string | null = null;
	private isLoading = false;
	private hostResizeObserver: ResizeObserver | null = null;
	private titleObserver: MutationObserver | null = null;

	constructor(leaf: WorkspaceLeaf, private getAuthorName: () => string) {
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
		this.root = createRoot(this.hostEl);
		this.render();
	}

	async onClose() {
		this.root?.unmount();
		this.root = null;
		this.hostResizeObserver?.disconnect();
		this.hostResizeObserver = null;
		this.titleObserver?.disconnect();
		this.titleObserver = null;
		this.hostEl = null;
		this.reactViewRef = createRef<DocxReactViewHandle>();
		this.buffer = null;
		this.error = null;
	}

	async onLoadFile(file: TFile) {
		this.isLoading = true;
		this.error = null;
		this.buffer = null;
		this.render();

		try {
			this.buffer = await this.app.vault.readBinary(file);
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
		this.buffer = null;
		this.error = null;
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

		return await this.reactViewRef.current?.save() ?? false;
	}

	private async saveFile(buffer: ArrayBuffer) {
		const file = this.file;
		if (!file) {
			throw new Error('No docx file is open.');
		}

		await this.app.vault.modifyBinary(file, buffer);
		this.buffer = buffer.slice(0);
		this.render();
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

	private render() {
		this.root?.render(
			<DocxReactView
				ref={this.reactViewRef}
				file={this.file}
				buffer={this.buffer}
				error={this.error}
				isLoading={this.isLoading}
				authorName={this.getAuthorName()}
				onSave={(buffer) => this.saveFile(buffer)}
			/>,
		);
	}
}
