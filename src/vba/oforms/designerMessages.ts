// The messages the form designer's webview posts to its host - the one
// canvas script (preview.ts) speaks to both the MSForms designer and the
// VB6 designer, so the protocol lives here, beside the script, once.

export type DesignerMessage =
	| { type: 'geometry'; name: string; left?: number; top?: number; width?: number; height?: number }
	| { type: 'geometryBatch'; anchor?: string; items: { name: string; left?: number; top?: number; width?: number; height?: number }[] }
	| { type: 'add'; container: string; controlKind: string; left: number; top: number }
	| { type: 'remove'; name: string }
	| { type: 'reparent'; name: string; container: string; left: number; top: number }
	| { type: 'setProp'; name: string; prop: string; value: string }
	| { type: 'openHandler'; name: string; event: string }
	| { type: 'markupEdit'; text: string }
	| { type: 'docUndo' }
	| { type: 'docRedo' }
	| { type: 'docSave' }
	| { type: 'formResize'; width: number; height: number }
	| { type: 'paste'; names: string[] }
	| { type: 'removeMany'; names: string[] }
	| { type: 'zOrder'; name: string; toFront: boolean }
	| { type: 'tabOrder'; container: string; names: string[] };

/** The messages that change the form: everything but navigation, the pane's text, and the document commands. */
export type GestureMessage = Exclude<DesignerMessage,
	{ type: 'openHandler' } | { type: 'markupEdit' } | { type: 'docUndo' } | { type: 'docRedo' } | { type: 'docSave' }>;
