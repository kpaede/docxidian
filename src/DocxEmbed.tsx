import { App, MarkdownPostProcessorContext, MarkdownRenderChild, TFile } from 'obsidian';
import { createRoot, Root } from 'react-dom/client';
import { DocxEditor } from '@eigenpal/docx-editor-react';
import { ensureEditorStyles } from './DocxReactView';

const DOCX_EMBED_SELECTOR = '.internal-embed[src], .internal-embed[data-src]';

function getEmbedLinkPath(embedEl: Element) {
	return embedEl.getAttribute('src') ?? embedEl.getAttribute('data-src') ?? '';
}

function stripSubpath(linkPath: string) {
	return linkPath.split('#')[0] ?? '';
}

function isDocxLink(linkPath: string) {
	return stripSubpath(linkPath).toLowerCase().endsWith('.docx');
}

function resolveDocxEmbed(app: App, linkPath: string, sourcePath: string) {
	const file = app.metadataCache.getFirstLinkpathDest(stripSubpath(linkPath), sourcePath);
	return file instanceof TFile && file.extension.toLowerCase() === 'docx' ? file : null;
}

function DocxEmbedPreview({ file, buffer }: { file: TFile; buffer: ArrayBuffer }) {
	return (
		<DocxEditor
			key={`${file.path}-${file.stat.mtime}`}
			documentBuffer={buffer}
			documentName={file.basename}
			readOnly
		/>
	);
}

class DocxEmbedRenderChild extends MarkdownRenderChild {
	private root: Root | null = null;
	private unloaded = false;

	constructor(
		containerEl: HTMLElement,
		private app: App,
		private file: TFile,
	) {
		super(containerEl);
	}

	onload() {
		void this.loadDocument();
	}

	private async loadDocument() {
		this.containerEl.empty();
		this.containerEl.addClass('docxidian-embed');
		this.containerEl.createDiv({ cls: 'docxidian-embed-loading', text: `Loading ${this.file.name}...` });

		try {
			ensureEditorStyles();
			const buffer = await this.app.vault.readBinary(this.file);
			if (this.unloaded) {
				return;
			}

			this.containerEl.empty();
			const hostEl = this.containerEl.createDiv({ cls: 'docxidian-embed-host' });
			this.root = createRoot(hostEl);
			this.root.render(<DocxEmbedPreview file={this.file} buffer={buffer} />);
		} catch (error) {
			if (this.unloaded) {
				return;
			}

			const message = error instanceof Error ? error.message : 'Unknown error';
			this.containerEl.empty();
			this.containerEl.createDiv({
				cls: 'docxidian-embed-error',
				text: `Could not render ${this.file.name}: ${message}`,
			});
		}
	}

	onunload() {
		this.unloaded = true;
		this.root?.unmount();
		this.root = null;
	}
}

class DocxEmbedScanChild extends MarkdownRenderChild {
	private timeouts: number[] = [];

	constructor(
		containerEl: HTMLElement,
		private app: App,
		private ctx: MarkdownPostProcessorContext,
	) {
		super(containerEl);
	}

	onload() {
		this.scan();
		this.timeouts.push(window.setTimeout(() => this.scan(), 0));
		this.timeouts.push(window.setTimeout(() => this.scan(), 100));
	}

	onunload() {
		for (const timeout of this.timeouts) {
			window.clearTimeout(timeout);
		}
		this.timeouts = [];
	}

	private scan() {
		renderDocxEmbeds(this.app, this.containerEl, this.ctx);
	}
}

function collectEmbedElements(el: HTMLElement) {
	const embeds = Array.from(el.querySelectorAll(DOCX_EMBED_SELECTOR));
	if (el.matches(DOCX_EMBED_SELECTOR)) {
		embeds.unshift(el);
	}

	return embeds;
}

export function renderDocxEmbeds(app: App, el: HTMLElement, ctx: MarkdownPostProcessorContext) {
	const embeds = collectEmbedElements(el);

	for (const embedEl of embeds) {
		if (!(embedEl instanceof HTMLElement) || embedEl.dataset.docxidianEmbed === 'true') {
			continue;
		}

		const linkPath = getEmbedLinkPath(embedEl);
		if (!isDocxLink(linkPath)) {
			continue;
		}

		const file = resolveDocxEmbed(app, linkPath, ctx.sourcePath);
		if (!file) {
			continue;
		}

		embedEl.dataset.docxidianEmbed = 'true';
		ctx.addChild(new DocxEmbedRenderChild(embedEl, app, file));
	}
}

export function processDocxEmbeds(app: App, el: HTMLElement, ctx: MarkdownPostProcessorContext) {
	ctx.addChild(new DocxEmbedScanChild(el, app, ctx));
}
