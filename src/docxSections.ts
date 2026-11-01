import type { DocxEditorRef } from '@eigenpal/docx-editor-react';

type WordDocument = NonNullable<ReturnType<DocxEditorRef['getDocument']>>['package']['document'];
type SectionProperties = NonNullable<WordDocument['finalSectionProperties']>;

// The editor keeps the final section properties twice. Page setup and ruler
// edits must update both copies, while preserving earlier sections and content.
export function updateFinalSectionProperties(document: WordDocument, changes: Partial<SectionProperties>): WordDocument {
	const properties = { ...document.finalSectionProperties, ...changes };
	const sections = document.sections;
	return {
		...document,
		finalSectionProperties: properties,
		...(sections?.length ? {
			sections: sections.map((section, index) => index === sections.length - 1
				? { ...section, properties: { ...section.properties, ...properties } }
				: section),
		} : {}),
	};
}
