// Shared with the editor's layout adapter. Text anchoring must not scroll the
// viewport underneath a ruler drag, whose coordinates are relative to it.
const draggingUntil = new WeakMap<HTMLElement, number>();

export function beginRulerDrag(container: HTMLElement) {
	draggingUntil.set(container, Infinity);
}

export function endRulerDrag(container: HTMLElement) {
	// Allow the final React update/layout to finish after mouseup. There are no
	// timers or delayed DOM writes; expired entries are ignored automatically.
	draggingUntil.set(container, performance.now() + 100);
}

export function clearRulerDrag(container: HTMLElement) {
	draggingUntil.delete(container);
}

export function isRulerDragActive(container: HTMLElement | null | undefined) {
	return !!container && (draggingUntil.get(container) ?? 0) > performance.now();
}
