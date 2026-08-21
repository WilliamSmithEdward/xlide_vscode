// Shared curation for every reference generator: Excel, the generated hosts
// (Word, PowerPoint, Access) and the shared Office library.
//
// The pyVBAReference dumps are faithful to the COM type libraries, and the type
// libraries are lossy in two ways that matter to a language service:
//
//   1. Accessors that hand back a real object are declared `As Object`, because
//      the underlying IDispatch signature says so. `Chart.Axes` is typed Object,
//      not Axes, so a member chain dies at that hop.
//   2. Nothing carries prose. The Learn enrichment fills most of it, but a
//      member with no published page arrives with an empty description.
//
// Excel's hand-written model repaired case 1 for its own core by transcribing
// return types one member at a time. That does not scale to four libraries and
// 1,700 types, so the repair lives here instead: rules derived from the corpus,
// each one validated against Excel's transcribed answers (tests/vbaHostReturnCuration
// .test.ts), plus a small table for the facts no rule can reach.
//
// Nothing here invents a member or an object type. A candidate is accepted only
// when it names a class the corpus actually carries.

import fs from 'node:fs';
import path from 'node:path';

/** Dump kinds that carry members. Coclasses and their dispatch interfaces both do. */
export const CLASS_KINDS = new Set(['Class', 'Dispatch Interface', 'Interface']);

/**
 * Declared return types that name no type at all. `Variant` is deliberately
 * absent: a Variant return is a real, documented Variant, not a gap.
 */
const GENERIC_RETURNS = new Set(['Object', 'VARIANT', '']);

/** Members whose return IS the collection's element rather than the collection. */
const ELEMENT_ACCESSORS = new Set(['item', '_default', 'add', 'add2']);

/**
 * Return types that no type library or documentation page states, because they
 * are facts about the host's object graph rather than about the member: what a
 * given object's Parent is, which sheet is active, which of two types an
 * overloaded accessor hands back.
 *
 * Keyed `TypeName.MemberName`. The value is a bare type name in the same
 * library, or an array when the reference documents an either/or return - the
 * first entry is the one a chain follows, the rest widen member completion.
 *
 * Excel's entries are the ones its hand-written model already carried and the
 * test suite already pinned. The rest were read off the Office VBA reference on
 * learn.microsoft.com, one page per entry.
 */
export const CURATED_RETURNS = {
	Excel: {
		'Application.ActiveSheet': 'Worksheet',
		'Window.ActiveSheet': 'Worksheet',
		'Global.ActiveSheet': 'Worksheet',
		'Range.Parent': 'Worksheet',
		'Worksheet.Parent': 'Workbook',
		'Worksheets.Parent': 'Workbook',
		'Sheets.Parent': 'Workbook',
		'Workbook.Parent': 'Application',
		'Workbooks.Parent': 'Application',
		'Window.Parent': 'Application',
		'SlicerCache.Parent': 'Workbook',
		'SlicerCaches.Parent': 'Workbook',
		'DataBarBorder.Color': 'FormatColor',
		'Databar.AxisColor': 'FormatColor',
		'DrawingObjects.Item': 'Drawing',
		'PivotTable.ColumnFields': 'PivotFields',
		'PivotTable.DataFields': 'PivotFields',
		'PivotTable.HiddenFields': 'PivotFields',
		'PivotTable.PageFields': 'PivotFields',
		'PivotTable.RowFields': 'PivotFields',
		'PivotTable.VisibleFields': 'PivotFields',
		'Sheets.Item': ['Worksheet', 'Chart'],
		'Sheets.Add': ['Worksheet', 'Chart'],
		'Sheets.Add2': ['Worksheet', 'Chart'],
		// "Uses either LinearGradient or RectangularGradient" - excel.interior.gradient.
		'Interior.Gradient': ['LinearGradient', 'RectangularGradient'],
		// "Returns the Shape or Range object that is positioned at the specified
		// pair of screen coordinates" - excel.window.rangefrompoint.
		'Window.RangeFromPoint': ['Shape', 'Range'],
	},
	Word: {
		// "Returns a Template or Document object" - word.application.macrocontainer.
		'Application.MacroContainer': ['Template', 'Document'],
		'Global.MacroContainer': ['Template', 'Document'],
		// "Returns or sets a Template or Document object" - word.application.customizationcontext.
		'Application.CustomizationContext': ['Template', 'Document'],
		'Global.CustomizationContext': ['Template', 'Document'],
		// "can return Document, Template, or Application" - word.keybinding.context.
		'KeyBinding.Context': ['Document', 'Template', 'Application'],
		'KeyBindings.Context': ['Document', 'Template', 'Application'],
		'KeysBoundTo.Context': ['Document', 'Template', 'Application'],
		// "Returns the Range or Shape object ... at the point" - word.window.rangefrompoint.
		'Window.RangeFromPoint': ['Range', 'Shape'],
	},
	PowerPoint: {},
	Access: {
		// AccessProperty is a hidden member of the Access object model, so no
		// reference page names it (activex-access typings; Access World Forums
		// thread 132943, which prints TypeName of the returned object).
		'Properties.Item': 'AccessProperty',
	},
	Office: {
		// "Returns an ODSOColumns object" / "Returns an ODSOFilters collection" -
		// office.officedatasourceobject.columns and .filters.
		'OfficeDataSourceObject.Columns': 'ODSOColumns',
		'OfficeDataSourceObject.Filters': 'ODSOFilters',
		// The Office Assistant objects predate the modern reference; the types
		// come from the Office PIA and the archived Office 2003 VBA reference.
		'Balloon.Checkboxes': 'BalloonCheckboxes',
		'Balloon.Labels': 'BalloonLabels',
		// "Gets the Shape object" - the Office PIA links the name to its own
		// Shape class (activex-office typings agree).
		'Signature.SignatureLineShape': 'Shape',
	},
};

export function readDumps(jsonDir) {
	const dumps = new Map();
	for (const fileName of fs.readdirSync(jsonDir).sort()) {
		if (!fileName.endsWith('.json') || fileName === '_index.json') {
			continue;
		}
		try {
			const dump = JSON.parse(fs.readFileSync(path.join(jsonDir, fileName), 'utf8'));
			dumps.set(dump.name, dump);
		} catch {
			// A malformed dump stays out of the model rather than breaking it.
		}
	}
	return dumps;
}

/** Class names of a library, for the return-qualification fallback. */
export function classNamesIn(jsonDir) {
	const names = new Set();
	let dumps;
	try {
		dumps = readDumps(jsonDir);
	} catch {
		return names;   // library not present locally: those returns stay bare
	}
	for (const dump of dumps.values()) {
		if (CLASS_KINDS.has(dump.kind)) {
			names.add(dump.name);
		}
	}
	return names;
}

/**
 * Every description string a library publishes, per Office application. Used to
 * prove that a sentence naming another application was cross-published rather
 * than written about it.
 */
export function descriptionIndex(referenceRoot) {
	const byApp = new Map();
	for (const [dir, app] of [['excel', 'Excel'], ['word', 'Word'], ['powerpoint', 'PowerPoint'], ['access', 'Access']]) {
		const texts = new Set();
		let dumps;
		try {
			dumps = readDumps(path.join(referenceRoot, dir, 'json'));
		} catch {
			byApp.set(app, texts);
			continue;
		}
		for (const dump of dumps.values()) {
			const push = (value) => {
				const text = collapseWhitespace(value);
				if (text) {
					texts.add(text);
				}
			};
			push(dump.description);
			push(dump.remarks);
			for (const member of [...(dump.properties ?? []), ...(dump.methods ?? []), ...(dump.constants ?? [])]) {
				push(member.description);
				for (const param of member.parameters ?? []) {
					push(param.description);
				}
			}
		}
		byApp.set(app, texts);
	}
	return byApp;
}

/**
 * The reference publishes one page per SHARED object under every host namespace
 * and never substitutes the application name, so PowerPoint's chart members
 * describe themselves as Word: "True if Microsoft Word plots data points from
 * last to first" is what a PowerPoint developer read on hover.
 *
 * A mention is rewritten only when the whole sentence appears verbatim in the
 * named application's own library, which is what proves it was cross-published.
 * A sentence written ABOUT another application - "an external Microsoft Excel
 * workbook", "pasting from Microsoft PowerPoint" - exists only in the library
 * that says it, so it is left exactly as the reference wrote it.
 */
export function localizeHostName(text, hostApp, index) {
	if (!text || !hostApp || !index) {
		return text;
	}
	let out = text;
	for (const [app, texts] of index) {
		if (app === hostApp || !out.includes(`Microsoft ${app}`) || !texts.has(text)) {
			continue;
		}
		out = out.split(`Microsoft ${app}`).join(`Microsoft ${hostApp}`);
	}
	return out;
}

export function collapseWhitespace(text) {
	if (typeof text !== 'string') {
		return undefined;
	}
	const line = text.replace(/\s+/g, ' ').trim();
	return line.length > 0 ? line : undefined;
}

function truncate(text, max) {
	if (!text || text.length <= max) {
		return text;
	}
	return `${text.slice(0, max - 1)}…`;
}

/**
 * Singular forms of a collection name, most specific first. Only used to name a
 * collection's element type, and only accepted when the result is a real class,
 * so an over-eager trim (`Axis` from `Axis`) simply finds nothing.
 */
export function singularCandidates(name) {
	const out = [];
	if (/(?:s|x|z|ch|sh)es$/i.test(name)) {
		out.push(name.slice(0, -2));            // TextBoxes -> TextBox
	}
	if (/[^aeiou]ies$/i.test(name)) {
		out.push(`${name.slice(0, -3)}y`);      // Entries -> Entry
	}
	if (/s$/i.test(name) && !/ss$/i.test(name)) {
		out.push(name.slice(0, -1));            // Worksheets -> Worksheet
	}
	return out;
}

/** The element type a collection's own description names. */
function describedElementType(dump) {
	const text = String(dump?.description ?? '');
	return text.match(/\bcollection of (?:all )?(?:the )?([A-Z][A-Za-z0-9_]*) objects?\b/)?.[1]
		?? text.match(/\bcollection of (?:all )?(?:the )?([A-Z][A-Za-z0-9_]*)s\b/)?.[1];
}

/**
 * Type names a member's own description states. The reference prose is
 * formulaic enough to read: an aside names the concrete type behind an `Object`
 * return ("...the fields (a PivotFields object)"), a lead sentence names it
 * outright ("Returns a Font object..."), and the access tail repeats it
 * ("...Read-only Range."). The LAST aside wins, because a description that
 * lists both an element and its collection names the collection second.
 */
function describedReturnTypes(description) {
	const text = String(description ?? '');
	const out = [];
	const asides = [...text.matchAll(/\(an? ([A-Z][A-Za-z0-9_]*) (?:object|collection)\)/g)];
	if (asides.length > 0) {
		out.push(asides[asides.length - 1][1]);
	}
	const lead = text.match(
		/\b(?:Returns?|Gets|Sets|Returns or sets|Gets or sets) (?:a|an|the) (?:new )?([A-Z][A-Za-z0-9_]*) (?:object|collection)\b/,
	);
	if (lead) {
		out.push(lead[1]);
	}
	const tail = text.match(/\b(?:Read[- ]only|Read\/write|Write[- ]only) ([A-Z][A-Za-z0-9_]*)\.\s*$/);
	if (tail) {
		out.push(tail[1]);
	}
	return out;
}

/**
 * Creates the per-library curator. `prefix` is the model's namespace ("Excel"),
 * `foreignClasses` maps a bare class name to the namespace that owns it, for
 * the shared Office types every host library returns.
 */
export function createCurator({ dumps, prefix, foreignClasses = new Map() }) {
	const classNames = new Set(
		[...dumps.values()].filter((dump) => CLASS_KINDS.has(dump.kind)).map((dump) => dump.name),
	);
	const curated = CURATED_RETURNS[prefix] ?? {};

	/** The namespace that owns a bare class name, or undefined when nothing does. */
	function namespaceOf(bare) {
		if (classNames.has(bare)) {
			return prefix;
		}
		return foreignClasses.get(bare);
	}

	/**
	 * A member's declared return, qualified when it names a modelled class and
	 * dropped otherwise. `returns` is the chaining type and nothing else: the
	 * receiver resolver reads an absent one as "this hop states no object", which
	 * is exactly what `As Long` and `As Variant` mean to a member chain. The
	 * declared type is not lost - it is what `signature` renders.
	 */
	function qualifyReturn(declared) {
		const bare = collapseWhitespace(declared);
		if (!bare) {
			return undefined;
		}
		const namespace = namespaceOf(bare);
		return namespace ? `${namespace}.${bare}` : undefined;
	}

	/**
	 * What a member hands back: `{ returns, returnsAnyOf }`, repairing the `As
	 * Object` returns the type library cannot express. Rules run most-specific
	 * first and every candidate must name a class the corpus carries, so an
	 * unrecognised one leaves the declared `Object` in place rather than
	 * guessing. `returnsAnyOf` appears only for a documented either/or.
	 */
	function resolveReturn(ownerName, raw, kind) {
		const declared = collapseWhitespace(kind === 'property' ? raw.type : raw.returns);
		if (!GENERIC_RETURNS.has(declared ?? '')) {
			return { returns: qualifyReturn(declared) };
		}
		const documented = curated[`${ownerName}.${String(raw.name ?? '')}`];
		if (Array.isArray(documented)) {
			const qualified = documented.map(namespaced).filter(Boolean);
			if (qualified.length > 1) {
				// No single `returns`: a caller asking for one type must get
				// nothing rather than the first of several.
				return { returnsAnyOf: qualified };
			}
			if (qualified.length === 1) {
				return { returns: qualified[0] };
			}
		}
		for (const candidate of repairCandidates(ownerName, raw, documented)) {
			const qualified = namespaced(candidate);
			if (qualified) {
				return { returns: qualified };
			}
		}
		return { returns: qualifyReturn(declared) };
	}

	function namespaced(bare) {
		if (typeof bare !== 'string' || !bare) {
			return undefined;
		}
		const namespace = namespaceOf(bare);
		return namespace ? `${namespace}.${bare}` : undefined;
	}

	function repairCandidates(ownerName, raw, documented) {
		const name = String(raw.name ?? '');
		const out = [typeof documented === 'string' ? documented : undefined];
		if (ELEMENT_ACCESSORS.has(name.toLowerCase())) {
			out.push(describedElementType(dumps.get(ownerName)), ...singularCandidates(ownerName));
		}
		out.push(...describedReturnTypes(raw.description));
		if (name.toLowerCase() === 'duplicate') {
			out.push(ownerName);   // a clone hands back its own type
		}
		out.push(name);
		return out;
	}

	return { classNames, namespaceOf, qualifyReturn, resolveReturn };
}

/**
 * A method's call signature, synthesized from its parameter list when the corpus
 * carries none. Properties get none by design: `signature` is what signature
 * help offers and what call-arity checking measures, and a property's declared
 * type is neither. That fact travels as `declaredType` instead.
 */
export function memberSignature(raw, kind) {
	if (kind !== 'method') {
		return undefined;
	}
	const declared = collapseWhitespace(raw.returns);
	const signature = collapseWhitespace(raw.signature);
	if (signature) {
		return signature;
	}
	const params = (raw.parameters ?? []).map((param) => {
		const text = `${param.name}${param.type ? ` As ${param.type}` : ''}`;
		return param.optional ? `[${text}]` : text;
	});
	const call = `${raw.name}(${params.join(', ')})`;
	return declared && declared !== 'void' ? `${call} As ${declared}` : call;
}

/**
 * A member's documentation. `params` carries the reference's own per-parameter
 * prose, which signature help shows against the active parameter; the generated
 * hosts used to drop it on the floor.
 */
export function memberDoc(raw, maxSummary = 300, hostApp = undefined, index = undefined) {
	const summary = truncate(localizeHostName(collapseWhitespace(raw.description), hostApp, index), maxSummary);
	const params = (raw.parameters ?? [])
		.map((param) => ({
			name: String(param.name ?? ''),
			text: localizeHostName(collapseWhitespace(param.description), hostApp, index) ?? '',
			...(collapseWhitespace(param.type) ? { type: collapseWhitespace(param.type) } : {}),
		}))
		.filter((param) => param.name && (param.text || param.type));
	if (!summary && params.length === 0) {
		return undefined;
	}
	return {
		...(summary ? { summary } : {}),
		params,
		source: 'external',
	};
}

/**
 * A type's documentation: the reference's description, its remarks, and its
 * worked example. Only the description was ever read before, so `Dim s As
 * InlineShape` hovered with a bare type name and nothing else.
 */
export function typeDoc(dump, maxSummary = 300, hostApp = undefined, index = undefined) {
	const summary = truncate(localizeHostName(collapseWhitespace(dump.description), hostApp, index), maxSummary);
	const remarks = truncate(localizeHostName(collapseWhitespace(dump.remarks), hostApp, index), 600);
	const example = typeof dump.example === 'string' && dump.example.trim() ? dump.example.trim() : undefined;
	if (!summary && !remarks && !example) {
		return undefined;
	}
	return {
		...(summary ? { summary } : {}),
		...(remarks ? { remarks } : {}),
		...(example ? { example } : {}),
		params: [],
		source: 'external',
	};
}

/**
 * The type a property declares. Distinct from `returns`, which names only a
 * modelled object type for chaining: `Range.Value` declares Variant, chains
 * nowhere, and should still hover as `Value As Variant`.
 */
export function declaredType(raw, kind) {
	return kind === 'property' ? collapseWhitespace(raw.type) : undefined;
}

/** The read/write contract of a property, as the type library states it. */
export function memberAccess(raw, kind) {
	if (kind !== 'property') {
		return undefined;
	}
	const access = collapseWhitespace(raw.access);
	return access === 'read-only' || access === 'read/write' || access === 'write-only'
		? access
		: undefined;
}
