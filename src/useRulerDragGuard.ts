import { useEffect, type MutableRefObject } from 'react';
import { beginRulerDrag, clearRulerDrag, endRulerDrag } from './rulerDragState';

export function useRulerDragGuard(editorClassName: MutableRefObject<string>) {
	useEffect(() => {
		let scrollContainer: HTMLElement | null = null;
		let dragging = false;
		const start = (event: MouseEvent) => {
			if (event.button !== 0 || !(event.target instanceof Element)) return;
			const ruler = event.target.closest('.docx-vertical-ruler, .docx-horizontal-ruler');
			const editor = document.querySelector(`.${editorClassName.current}`);
			if (!ruler || !editor?.contains(ruler)) return;
			const next = editor.querySelector<HTMLElement>('.docx-editor__scroll-container');
			if (!next) return;
			if (scrollContainer && scrollContainer !== next) clearRulerDrag(scrollContainer);
			scrollContainer = next;
			dragging = true;
			beginRulerDrag(next);
		};
		const finish = () => {
			if (scrollContainer && dragging) endRulerDrag(scrollContainer);
			dragging = false;
		};
		document.addEventListener('mousedown', start, true);
		document.addEventListener('mouseup', finish, true);
		window.addEventListener('blur', finish);
		return () => {
			document.removeEventListener('mousedown', start, true);
			document.removeEventListener('mouseup', finish, true);
			window.removeEventListener('blur', finish);
			if (scrollContainer) clearRulerDrag(scrollContainer);
		};
	}, [editorClassName]);
}
