/**
 * Edit-script planning for in-place OOXML saves.
 *
 * A plan compares each part's anchored unit sequence with the current model
 * sequence. Unchanged units stay inside the entry's carried bytes, changed
 * units replace their anchor span, new units insert at a neighbouring anchor
 * boundary, and removed units delete their anchor span. Table cell geometry is
 * refused because row or column changes cannot be expressed as independent
 * span operations, and a caller-supplied classifier can refuse a coupled
 * construct that overlaps an edited span.
 *
 * Every plan either covers each source byte exactly once, outside a
 * `replace`/`delete` range or inside one, or it carries a refusal and no
 * operations. The planner never reads part bytes from the archive; providers
 * supply rendered unit bytes through {@link SplicePlanInput.renderUnit}.
 * @module @forwork/office-exchange/source/splice
 */

import type { UnsupportedFeatureKind } from '../types.ts'
import type { SourcePackage, ZipEntryRecord } from '../package/types.ts'
import {
  CDATA_CLOSE,
  CDATA_OPEN,
  COMMENT_CLOSE,
  COMMENT_OPEN,
  PROCESSING_CLOSE,
  declarationEnd,
  decodeBytes,
  findSequence,
  findTagEnd,
  indexOfByte,
  parseStartTag,
  readName,
  splitQName,
  startsWith,
} from './xml-scan.ts'
import type { ElementSpan } from './xml-spans.ts'
import { sourceDispositions } from './save-disposition.ts'
import type { OfficeSourceState, UnitAnchor, UnitAnchorKind } from './source-state.ts'

/** One current editable unit in model order. */
export interface CurrentUnit {
  /** Unit kind the anchor names. */
  readonly kind: UnitAnchorKind
  /** Decoded part name the unit belongs to. */
  readonly part: string
  /** Model identity: `paragraphId`, `${tableId}/${row}/${cell}`, or `${sheetId}!${row}:${column}`. */
  readonly key: string
  /** Current canonical unit digest. */
  readonly digest: string
}

/** An auxiliary element the provider appends before a part's closing tag. */
export interface SpliceAppend {
  /** Existing source part the bytes append to. */
  readonly part: string
  /** Byte offset before which the bytes are inserted. */
  readonly at: number
  /** Complete serialized element bytes. */
  readonly bytes: Uint8Array
}

/** A package entry the caller adds or removes outside the span plan. */
export interface SpliceEntryChange {
  /** ZIP entry name from the central directory, or the new entry name. */
  readonly name: string
  /** Why the entry is added or removed. */
  readonly reason: string
}

/** One byte-affecting edit the classifier may refuse. */
export interface SpliceEdit {
  /** Operation the edit produces. */
  readonly kind: 'replace' | 'insert' | 'delete'
  /** Decoded part name. */
  readonly part: string
  /** Unit kind the edit concerns. */
  readonly unitKind: UnitAnchorKind
  /** Unit key the edit concerns. */
  readonly key: string
  /** Inclusive source byte offset; equal to {@link end} for an insert. */
  readonly start: number
  /** Exclusive source byte offset; equal to {@link start} for an insert. */
  readonly end: number
}

/** Caller-supplied coupled-construct classifier over one planned edit. */
export type CoupledConstructClassifier = (edit: SpliceEdit) => ConstructIntersection | undefined

/** Inputs the writer supplies to {@link buildSplicePlan}. */
export interface SplicePlanInput {
  /** Complete source package as read at the expected version. */
  readonly source: SourcePackage
  /** Validated sidecar describing the source units. */
  readonly sourceState: OfficeSourceState
  /**
   * Explicit anchors. Required for `sheet-summary` state; otherwise the
   * per-unit anchors from {@link sourceState} are used.
   */
  readonly anchors?: readonly UnitAnchor[]
  /** Current units in model order, grouped by part when the caller projects one part's sequence. */
  readonly units: readonly CurrentUnit[]
  /** Spans recorded by `xml-spans` for the source parts. */
  readonly spans: readonly ElementSpan[]
  /** Render one current unit's replacement bytes. */
  readonly renderUnit: (unit: CurrentUnit) => Uint8Array
  /** Refuse an edit that overlaps a coupled construct. */
  readonly classifyCoupled?: CoupledConstructClassifier
  /** Auxiliary elements appended before a source part's closing tag. */
  readonly appends?: readonly SpliceAppend[]
  /** New package entries outside the span plan. */
  readonly addedEntries?: readonly SpliceEntryChange[]
  /** Removed package entries outside the span plan. */
  readonly removedEntries?: readonly SpliceEntryChange[]
}

/** A typed construct intersection that prevents the plan from completing. */
export interface ConstructIntersection {
  /** Construct class from the shared unsupported-feature vocabulary. */
  readonly kind: UnsupportedFeatureKind
  /** Actionable description naming the construct and its location. */
  readonly detail: string
  /** Part the construct overlaps. */
  readonly part: string
  /** Source byte offset or anchor-relative position of the construct. */
  readonly start: number
}

/** What happened to one package entry. */
export interface PartDisposition {
  /** ZIP entry name as stored in the central directory. */
  readonly name: string
  /** Entry outcome. */
  readonly status: 'verbatim' | 'spliced' | 'added' | 'removed'
}

/** No byte change to a source entry. */
export interface CarryOperation {
  /** Operation discriminant. */
  readonly kind: 'carry'
  /** Decoded part name. */
  readonly part: string
  /** ZIP entry name carried unchanged. */
  readonly entry: string
}

/** Replace one anchored unit's source span. */
export interface ReplaceOperation {
  /** Operation discriminant. */
  readonly kind: 'replace'
  /** Decoded part name. */
  readonly part: string
  /** Inclusive source byte offset. */
  readonly start: number
  /** Exclusive source byte offset. */
  readonly end: number
  /** Rendered unit bytes. */
  readonly bytes: Uint8Array
}

/** Insert a new unit at an anchor boundary. */
export interface InsertOperation {
  /** Operation discriminant. */
  readonly kind: 'insert'
  /** Decoded part name. */
  readonly part: string
  /** Source byte offset before which the bytes are inserted. */
  readonly at: number
  /** Rendered unit bytes. */
  readonly bytes: Uint8Array
}

/** Delete one anchored unit's source span. */
export interface DeleteOperation {
  /** Operation discriminant. */
  readonly kind: 'delete'
  /** Decoded part name. */
  readonly part: string
  /** Inclusive source byte offset. */
  readonly start: number
  /** Exclusive source byte offset. */
  readonly end: number
}

/** Append auxiliary content before a part's closing tag. */
export interface AppendOperation {
  /** Operation discriminant. */
  readonly kind: 'append'
  /** Decoded part name. */
  readonly part: string
  /** Source byte offset before which the bytes are inserted. */
  readonly at: number
  /** Complete serialized element bytes. */
  readonly bytes: Uint8Array
}

/** One ordered byte-range operation in a splice plan. */
export type SpliceOperation =
  | CarryOperation
  | ReplaceOperation
  | InsertOperation
  | DeleteOperation
  | AppendOperation

/** One splice operation that changes source bytes; carried entries are excluded. */
export type ByteSpliceOperation = Exclude<SpliceOperation, CarryOperation>

/** The writer's ordered edit script and per-entry dispositions. */
export interface SplicePlan {
  /** Operations sorted by part name and source position; non-overlapping per part. */
  readonly operations: readonly SpliceOperation[]
  /** Parts named by span operations or auxiliary appends, sorted ASCII. */
  readonly touchedParts: readonly string[]
  /** One disposition per file entry, sorted by entry name. */
  readonly entries: readonly PartDisposition[]
  /** Present when the plan cannot be expressed as span operations. */
  readonly refusal?: ConstructIntersection
}

/** A source property the renderer did not reproduce and the caller must merge. */
export interface IndependentProperty {
  /** Property container that owns the property: a Word property element, the sheet cell root, or the unit root. */
  readonly container: 'pPr' | 'rPr' | 'tcPr' | 'trPr' | 'c' | 'root'
  /** Attribute or child element qualified name as written in the source. */
  readonly name: string
  /** Verbatim source attribute or child element bytes. */
  readonly xml: string
  /** Inclusive byte offset in the source span. */
  readonly start: number
  /** Exclusive byte offset in the source span. */
  readonly end: number
}

/** An unreproduced source construct outside the independent-property whitelist. */
export interface CoupledConstructCandidate {
  /** Attribute or element qualified name as written in the source. */
  readonly name: string
  /** Verbatim source text of the candidate. */
  readonly xml: string
  /** Inclusive byte offset in the source span. */
  readonly start: number
  /** Exclusive byte offset in the source span. */
  readonly end: number
  /** Why the candidate cannot be merged as an independent property. */
  readonly detail: string
}

/** The result of comparing a source unit span with its rendered replacement. */
export interface IndependentPropertyClassification {
  /** Whitelisted source properties the renderer did not reproduce. */
  readonly merge: readonly IndependentProperty[]
  /** Word root metadata attributes the caller re-emits on the rendered root tag. */
  readonly rootAttributes: readonly IndependentProperty[]
  /** Unreproduced constructs outside the whitelist; never merged. */
  readonly coupled: readonly CoupledConstructCandidate[]
}

/** Property container local names whose children are carried instead of refused. */
type PropertyContainer = 'pPr' | 'rPr' | 'tcPr' | 'trPr'

interface ParsedAttribute {
  readonly qname: string
  readonly namespace: string
  readonly localName: string
  readonly value: string
  readonly start: number
  readonly end: number
}

interface ParsedElement {
  readonly qname: string
  readonly namespace: string
  readonly localName: string
  readonly start: number
  end: number
  readonly attributes: readonly ParsedAttribute[]
  readonly children: ParsedElement[]
}

interface OpenElement {
  readonly element: ParsedElement
  readonly bindings: ReadonlyMap<string, string>
}

type ParseResult =
  | { readonly ok: true; readonly root: ParsedElement | undefined }
  | { readonly ok: false; readonly detail: string }

interface PropertyRecord {
  readonly container: PropertyContainer
  readonly element: ParsedElement
}

interface CandidateRecord {
  readonly element: ParsedElement
  readonly detail: string
}

interface StructureScan {
  readonly properties: readonly PropertyRecord[]
  readonly candidates: readonly CandidateRecord[]
}

const WORD_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
])

const SHEET_NAMESPACES: ReadonlySet<string> = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
])

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace'

const WORD14_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordml'

const ROOT_EDITOR_ID_NAMES: ReadonlySet<string> = new Set(['paraId', 'textId'])

const WORD_ROOT_ATTRIBUTE_NAMESPACES: ReadonlySet<string> = new Set([...WORD_NAMESPACES, 'w', ''])

const WORD14_ROOT_ATTRIBUTE_NAMESPACES: ReadonlySet<string> = new Set([WORD14_NAMESPACE, 'w14'])

const PROPERTY_CONTAINERS: ReadonlySet<string> = new Set(['pPr', 'rPr', 'tcPr', 'trPr'])

const REVISION_CHANGE_ELEMENTS: ReadonlySet<string> = new Set([
  'pPrChange',
  'rPrChange',
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'sectPrChange',
])

const WORD_CONTENT_ELEMENTS: ReadonlySet<string> = new Set([
  'p',
  'r',
  't',
  'br',
  'tab',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'sym',
  'drawing',
  'pict',
  'object',
  'hyperlink',
  'smartTag',
  'customXml',
])

const SHEET_CONTENT_ELEMENTS: ReadonlySet<string> = new Set(['c', 'v', 'f', 'is', 't'])

const SHARED_FORMULA_TYPES: ReadonlySet<string> = new Set(['shared', 'array'])

/**
 * Build the ordered splice plan for one save.
 *
 * The plan is computed without decoding source parts. A refusal returns no
 * operations, no touched parts, and no dispositions.
 * @param input - Source package, sidecar anchors, current units, spans, and the provider's unit renderer.
 * @returns The plan, or a refusal with zero operations.
 */
export function buildSplicePlan(input: SplicePlanInput): SplicePlan {
  const entriesByName = sourceEntryMap(input.source)
  const anchors = input.anchors ?? (input.sourceState.anchorMode === 'per-unit' ? input.sourceState.anchors : undefined)
  if (anchors === undefined) {
    return refusalPlan({
      kind: 'other',
      detail: 'sheet-summary state requires rebuilt per-unit anchors before splicing',
      part: input.sourceState.sourceParts[0] ?? '',
      start: 0,
    })
  }
  const inputRefusal = validateInput(input, entriesByName, anchors)
  if (inputRefusal !== undefined) return refusalPlan(inputRefusal)

  const anchorsByPart = groupAnchors(anchors)
  const unitsByPart = groupUnits(input.units)
  const removedNames = new Set((input.removedEntries ?? []).map(entry => entry.name))

  const parts = new Set<string>([
    ...anchorsByPart.keys(),
    ...unitsByPart.keys(),
    ...(input.appends ?? []).map(append => append.part),
  ])
  const sortedParts = [...parts].sort(compareAscii)

  for (const part of sortedParts) {
    const partAnchors = anchorsByPart.get(part) ?? []
    const partUnits = unitsByPart.get(part) ?? []
    const geometry = tableGeometryRefusal(part, partAnchors, partUnits)
    if (geometry !== undefined) return refusalPlan(geometry)
    const reorder = reorderRefusal(part, partAnchors, partUnits)
    if (reorder !== undefined) return refusalPlan(reorder)
  }

  const operations: SpliceOperation[] = []
  const edits: SpliceEdit[] = []

  for (const part of sortedParts) {
    const partAnchors = anchorsByPart.get(part) ?? []
    const partUnits = unitsByPart.get(part) ?? []
    const diff = diffPart(part, partAnchors, partUnits, input.renderUnit)
    if (!diff.ok) return refusalPlan(diff.refusal)
    operations.push(...diff.operations)
    edits.push(...diff.edits)
  }

  for (const append of input.appends ?? []) {
    operations.push({ kind: 'append', part: append.part, at: append.at, bytes: append.bytes })
  }

  const classifier = input.classifyCoupled
  if (classifier !== undefined) {
    for (const edit of sortEdits(edits)) {
      const refusal = classifier(edit)
      if (refusal !== undefined) return refusalPlan(refusal)
    }
  }

  for (const entry of input.source.entries) {
    if (entry.name.endsWith('/') || removedNames.has(entry.name)) continue
    if (operations.some(operation => operation.part === entry.name || (operation.kind === 'carry' && operation.entry === entry.name))) {
      continue
    }
    operations.push({ kind: 'carry', part: entry.name, entry: entry.name })
  }

  const sortedOperations = sortOperations(operations)
  const touchedParts = touchedPartsOf(sortedOperations)
  const entries = buildDispositions(input, sortedOperations, removedNames)

  return { operations: sortedOperations, touchedParts, entries }
}

/**
 * Compare a source unit span with its rendered replacement.
 *
 * Only the independent whitelist is returned for merging: a spreadsheet cell's
 * unprefixed `s` style index, direct children of `w:tcPr` and `w:trPr`, and
 * unmatched children of `w:pPr` and `w:rPr`. Word root metadata attributes
 * (`w:rsid*`, `w14:paraId`, and `w14:textId`) are returned separately for the
 * caller to re-emit on the rendered root tag. Everything else the renderer did
 * not reproduce is a candidate coupled construct, including revision-change
 * metadata, tracked-change wrappers, comment ranges, content controls, field
 * markup, and shared or array formulas.
 *
 * Text content differences are the edit itself and are ignored. Attribute
 * presence, not value, decides whether a root attribute was reproduced, so a
 * mapped value change is not reported as an unmapped property.
 * @param sourceSpanBytes - Uncompressed source bytes of the unit's anchor span.
 * @param renderedUnitBytes - Bytes the provider rendered for the same unit.
 * @returns Whitelisted properties to merge and unreproduced coupled candidates.
 */
export function classifyIndependentProperties(
  sourceSpanBytes: Uint8Array,
  renderedUnitBytes: Uint8Array,
): IndependentPropertyClassification {
  const source = parseUnitXml(sourceSpanBytes)
  if (!source.ok) {
    return {
      merge: [],
      rootAttributes: [],
      coupled: [{
        name: 'xml',
        xml: decodeBytes(sourceSpanBytes),
        start: 0,
        end: sourceSpanBytes.byteLength,
        detail: `source span is not well-formed XML: ${source.detail}`,
      }],
    }
  }
  const rendered = parseUnitXml(renderedUnitBytes)
  const renderedRoot = rendered.ok ? rendered.root : undefined
  const renderedScan = renderedRoot === undefined ? emptyScan() : scanStructure(renderedRoot)
  const sourceRoot = source.root
  if (sourceRoot === undefined) {
    const coupled: CoupledConstructCandidate[] = []
    if (!rendered.ok) {
      coupled.push(renderedParseCandidate(renderedUnitBytes, rendered.detail))
    }
    return { merge: [], rootAttributes: [], coupled }
  }

  const merge: IndependentProperty[] = []
  const rootAttributes: IndependentProperty[] = []
  const coupled: CoupledConstructCandidate[] = []

  const sourceScan = scanStructure(sourceRoot)
  mergeRootStyleAttribute(sourceSpanBytes, sourceRoot, renderedRoot, merge, rootAttributes, coupled)
  if (renderedRoot !== undefined
    && (renderedRoot.namespace !== sourceRoot.namespace || renderedRoot.localName !== sourceRoot.localName)) {
    coupled.push({
      name: sourceRoot.qname,
      xml: sliceElement(sourceSpanBytes, sourceRoot),
      start: sourceRoot.start,
      end: sourceRoot.end,
      detail: `rendered unit root <${renderedRoot.qname}> does not match source root <${sourceRoot.qname}>`,
    })
  }

  const renderedProperties = countKeys(renderedScan.properties, record => propertyKey(record))
  const renderedCandidates = countKeys(renderedScan.candidates, record => elementKey(record.element))

  for (const record of sourceScan.properties) {
    const key = propertyKey(record)
    if (consumeKey(renderedProperties, key)) continue
    merge.push({
      container: record.container,
      name: record.element.qname,
      xml: sliceElement(sourceSpanBytes, record.element),
      start: record.element.start,
      end: record.element.end,
    })
  }

  for (const record of sourceScan.candidates) {
    if (consumeKey(renderedCandidates, elementKey(record.element))) continue
    coupled.push({
      name: record.element.qname,
      xml: sliceElement(sourceSpanBytes, record.element),
      start: record.element.start,
      end: record.element.end,
      detail: record.detail,
    })
  }

  if (!rendered.ok) {
    coupled.push(renderedParseCandidate(renderedUnitBytes, rendered.detail))
  }

  merge.sort((left, right) => left.start - right.start)
  rootAttributes.sort((left, right) => left.start - right.start)
  coupled.sort((left, right) => left.start - right.start)
  return { merge, rootAttributes, coupled }
}

type PartDiffResult =
  | { readonly ok: true; readonly operations: SpliceOperation[]; readonly edits: SpliceEdit[] }
  | { readonly ok: false; readonly refusal: ConstructIntersection }

function diffPart(
  part: string,
  anchors: readonly UnitAnchor[],
  units: readonly CurrentUnit[],
  renderUnit: (unit: CurrentUnit) => Uint8Array,
): PartDiffResult {
  const operations: SpliceOperation[] = []
  const edits: SpliceEdit[] = []
  const anchorByIdentity = new Map<string, UnitAnchor>()
  for (const anchor of anchors) anchorByIdentity.set(identityOf(anchor), anchor)
  const unitByIdentity = new Map<string, CurrentUnit>()
  for (const unit of units) unitByIdentity.set(identityOf(unit), unit)

  for (const anchor of anchors) {
    const unit = unitByIdentity.get(identityOf(anchor))
    if (unit === undefined) {
      operations.push({ kind: 'delete', part, start: anchor.start, end: anchor.end })
      edits.push({
        kind: 'delete',
        part,
        unitKind: anchor.kind,
        key: anchor.key,
        start: anchor.start,
        end: anchor.end,
      })
      continue
    }
    if (unit.digest === anchor.digest) continue
    operations.push({ kind: 'replace', part, start: anchor.start, end: anchor.end, bytes: renderUnit(unit) })
    edits.push({
      kind: 'replace',
      part,
      unitKind: unit.kind,
      key: unit.key,
      start: anchor.start,
      end: anchor.end,
    })
  }

  const insertGroups = new Map<number, Uint8Array[]>()
  const insertionOffsets = insertionOffsetsOf(units, anchorByIdentity)
  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index]
    if (unit === undefined || anchorByIdentity.has(identityOf(unit))) continue
    const at = insertionOffsets[index]
    if (at === undefined) {
      return {
        ok: false,
        refusal: invalidInput(
          `unit ${unit.key} in ${part} has no neighbouring source anchor`,
          part,
          0,
        ),
      }
    }
    const bytes = renderUnit(unit)
    const group = insertGroups.get(at)
    if (group === undefined) insertGroups.set(at, [bytes])
    else group.push(bytes)
    edits.push({
      kind: 'insert',
      part,
      unitKind: unit.kind,
      key: unit.key,
      start: at,
      end: at,
    })
  }

  for (const [at, chunks] of insertGroups) {
    operations.push({ kind: 'insert', part, at, bytes: concatBytes(chunks) })
  }

  return { ok: true, operations, edits }
}

/**
 * Resolve each unit's nearest anchored predecessor or successor offset.
 *
 * A forward pass records the end of the closest anchored unit before each
 * index; a backward pass fills the remaining entries with the start of the
 * closest anchored unit after it.
 * @param units - Current units in model order.
 * @param anchorByIdentity - Source anchors keyed by identity.
 * @returns One insertion offset per unit, or `undefined` when no anchor neighbours it.
 */
function insertionOffsetsOf(
  units: readonly CurrentUnit[],
  anchorByIdentity: ReadonlyMap<string, UnitAnchor>,
): readonly (number | undefined)[] {
  const offsets = new Array<number | undefined>(units.length)
  let predecessor: number | undefined
  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index]
    /* v8 ignore next -- The loop index stays inside the unit array. */
    if (unit === undefined) continue
    const anchor = anchorByIdentity.get(identityOf(unit))
    if (anchor !== undefined) predecessor = anchor.end
    offsets[index] = predecessor
  }
  let successor: number | undefined
  for (let index = units.length - 1; index >= 0; index -= 1) {
    const unit = units[index]
    /* v8 ignore next -- The loop index stays inside the unit array. */
    if (unit === undefined) continue
    const anchor = anchorByIdentity.get(identityOf(unit))
    if (anchor !== undefined) successor = anchor.start
    else if (offsets[index] === undefined) offsets[index] = successor
  }
  return offsets
}

function validateInput(
  input: SplicePlanInput,
  entriesByName: ReadonlyMap<string, ZipEntryRecord>,
  anchors: readonly UnitAnchor[],
): ConstructIntersection | undefined {
  const spanKeys = new Set(input.spans.map(span => spanKey(span)))
  const anchorIdentities = new Set<string>()
  const anchorsByPart = new Map<string, UnitAnchor[]>()

  for (const anchor of anchors) {
    const entry = entriesByName.get(anchor.part)
    if (entry === undefined) {
      return invalidInput(`anchor ${anchor.key} names unknown part ${anchor.part}`, anchor.part, anchor.start)
    }
    if (anchor.start < 0 || anchor.end <= anchor.start || anchor.end > entry.uncompressedSize) {
      return invalidInput(`anchor ${anchor.key} has an invalid span in ${anchor.part}`, anchor.part, anchor.start)
    }
    const spanKeyValue = `${anchor.part}\u0000${anchor.kind}\u0000${String(anchor.start)}\u0000${String(anchor.end)}`
    if (!spanKeys.has(spanKeyValue)) {
      return invalidInput(
        `anchor ${anchor.key} has no matching source span in ${anchor.part} at ${String(anchor.start)}-${String(anchor.end)}`,
        anchor.part,
        anchor.start,
      )
    }
    const identity = identityOf(anchor)
    if (anchorIdentities.has(identity)) {
      return invalidInput(`anchor key ${anchor.key} repeats for ${anchor.kind} in ${anchor.part}`, anchor.part, anchor.start)
    }
    anchorIdentities.add(identity)
    const list = anchorsByPart.get(anchor.part)
    if (list === undefined) anchorsByPart.set(anchor.part, [anchor])
    else list.push(anchor)
  }

  for (const [part, partAnchors] of anchorsByPart) {
    const sorted = [...partAnchors].sort((left, right) => left.start - right.start || left.end - right.end)
    let previousEnd = -1
    for (const anchor of sorted) {
      if (anchor.start < previousEnd) {
        return invalidInput(`anchors overlap in ${part} at byte ${String(anchor.start)}`, part, anchor.start)
      }
      previousEnd = anchor.end
    }
  }

  const unitIdentities = new Set<string>()
  for (const unit of input.units) {
    if (!entriesByName.has(unit.part)) {
      return invalidInput(`unit ${unit.key} names unknown part ${unit.part}`, unit.part, 0)
    }
    const identity = identityOf(unit)
    if (unitIdentities.has(identity)) {
      return invalidInput(`current unit key ${unit.key} repeats for ${unit.kind} in ${unit.part}`, unit.part, 0)
    }
    unitIdentities.add(identity)
  }

  const addedNames = new Set<string>()
  for (const entry of input.addedEntries ?? []) {
    if (entriesByName.has(entry.name)) {
      return invalidInput(`added entry ${entry.name} already exists in the source`, entry.name, 0)
    }
    if (addedNames.has(entry.name)) {
      return invalidInput(`added entry ${entry.name} repeats`, entry.name, 0)
    }
    addedNames.add(entry.name)
  }

  const removedNames = new Set<string>()
  for (const entry of input.removedEntries ?? []) {
    if (!entriesByName.has(entry.name)) {
      return invalidInput(`removed entry ${entry.name} is not in the source`, entry.name, 0)
    }
    if (removedNames.has(entry.name)) {
      return invalidInput(`removed entry ${entry.name} repeats`, entry.name, 0)
    }
    removedNames.add(entry.name)
    if (anchors.some(anchor => anchor.part === entry.name) || input.units.some(unit => unit.part === entry.name)) {
      return invalidInput(`removed entry ${entry.name} still carries editable units`, entry.name, 0)
    }
  }

  for (const append of input.appends ?? []) {
    const entry = entriesByName.get(append.part)
    if (entry === undefined) {
      return invalidInput(`append names unknown part ${append.part}`, append.part, append.at)
    }
    if (removedNames.has(append.part)) {
      return invalidInput(`append names removed part ${append.part}`, append.part, append.at)
    }
    if (append.at < 0 || append.at > entry.uncompressedSize) {
      return invalidInput(`append at ${String(append.at)} is outside ${append.part}`, append.part, append.at)
    }
  }

  return undefined
}

function invalidInput(detail: string, part: string, start: number): ConstructIntersection {
  return { kind: 'other', detail, part, start }
}

function tableGeometryRefusal(
  part: string,
  anchors: readonly UnitAnchor[],
  units: readonly CurrentUnit[],
): ConstructIntersection | undefined {
  const sourceCells = anchors.filter(anchor => anchor.kind === 'tableCell')
  const currentCells = units.filter(unit => unit.kind === 'tableCell')
  if (sourceCells.length === 0 && currentCells.length === 0) return undefined

  const source = tableShapes(sourceCells.map(anchor => ({ key: anchor.key, start: anchor.start })))
  if ('detail' in source) return invalidInput(source.detail, part, source.start)
  const current = tableShapes(currentCells.map(unit => ({ key: unit.key, start: 0 })))
  if ('detail' in current) return invalidInput(current.detail, part, current.start)

  const tableIds = new Set([...source.tables.keys(), ...current.tables.keys()])
  for (const tableId of [...tableIds].sort(compareAscii)) {
    const before = source.tables.get(tableId)
    const after = current.tables.get(tableId)
    if (before === undefined || after === undefined) {
      return invalidInput(`table ${tableId} is added or removed, which changes table geometry`, part, 0)
    }
    if (before.rows.size !== after.rows.size || !sameRows(before, after)) {
      return invalidInput(`table ${tableId} has a row or column resize`, part, 0)
    }
  }
  return undefined
}

function reorderRefusal(
  part: string,
  anchors: readonly UnitAnchor[],
  units: readonly CurrentUnit[],
): ConstructIntersection | undefined {
  const currentIdentities = new Set(units.map(unit => identityOf(unit)))
  const anchorIdentities = new Set(anchors.map(anchor => identityOf(anchor)))
  const sourceOrder = anchors
    .filter(anchor => currentIdentities.has(identityOf(anchor)))
    .map(anchor => identityOf(anchor))
  const currentOrder = units
    .filter(unit => anchorIdentities.has(identityOf(unit)))
    .map(unit => identityOf(unit))
  for (let index = 0; index < sourceOrder.length; index += 1) {
    if (sourceOrder[index] !== currentOrder[index]) {
      return invalidInput(`unit reordering in ${part} cannot be expressed as span operations`, part, 0)
    }
  }
  return undefined
}

interface TableShape {
  readonly rows: ReadonlyMap<number, ReadonlySet<number>>
}

type TableShapes =
  | { readonly tables: ReadonlyMap<string, TableShape> }
  | { readonly detail: string; readonly start: number }

function tableShapes(cells: readonly { readonly key: string; readonly start: number }[]): TableShapes {
  const tables = new Map<string, Map<number, Set<number>>>()
  for (const cell of cells) {
    const parsed = parseTableCellKey(cell.key)
    if (parsed === undefined) {
      return { detail: `table cell key ${cell.key} is not tableId/row/cell`, start: cell.start }
    }
    const rows = tables.get(parsed.tableId) ?? new Map<number, Set<number>>()
    tables.set(parsed.tableId, rows)
    const tableCells = rows.get(parsed.row) ?? new Set<number>()
    rows.set(parsed.row, tableCells)
    tableCells.add(parsed.cell)
  }
  const shapes = new Map<string, TableShape>()
  for (const [tableId, rows] of tables) shapes.set(tableId, { rows })
  return { tables: shapes }
}

function parseTableCellKey(key: string): { readonly tableId: string; readonly row: number; readonly cell: number } | undefined {
  const cellSeparator = key.lastIndexOf('/')
  if (cellSeparator <= 0) return undefined
  const rowSeparator = key.lastIndexOf('/', cellSeparator - 1)
  if (rowSeparator <= 0) return undefined
  const row = Number(key.slice(rowSeparator + 1, cellSeparator))
  const cell = Number(key.slice(cellSeparator + 1))
  if (!Number.isSafeInteger(row) || row < 0) return undefined
  if (!Number.isSafeInteger(cell) || cell < 0) return undefined
  return { tableId: key.slice(0, rowSeparator), row, cell }
}

function sameRows(left: TableShape, right: TableShape): boolean {
  for (const [row, cells] of left.rows) {
    const other = right.rows.get(row)
    if (other === undefined || other.size !== cells.size) return false
    for (const cell of cells) {
      if (!other.has(cell)) return false
    }
  }
  return true
}

function sourceEntryMap(source: SourcePackage): ReadonlyMap<string, ZipEntryRecord> {
  const entries = new Map<string, ZipEntryRecord>()
  for (const entry of source.entries) {
    if (entry.name.endsWith('/')) continue
    entries.set(entry.name, entry)
  }
  return entries
}

function groupAnchors(anchors: readonly UnitAnchor[]): Map<string, UnitAnchor[]> {
  const grouped = new Map<string, UnitAnchor[]>()
  for (const anchor of anchors) {
    const list = grouped.get(anchor.part)
    if (list === undefined) grouped.set(anchor.part, [anchor])
    else list.push(anchor)
  }
  for (const list of grouped.values()) list.sort((left, right) => left.start - right.start)
  return grouped
}

function groupUnits(units: readonly CurrentUnit[]): Map<string, CurrentUnit[]> {
  const grouped = new Map<string, CurrentUnit[]>()
  for (const unit of units) {
    const list = grouped.get(unit.part)
    if (list === undefined) grouped.set(unit.part, [unit])
    else list.push(unit)
  }
  return grouped
}

function identityOf(unit: { readonly kind: UnitAnchorKind; readonly part: string; readonly key: string }): string {
  return `${unit.part}\u0000${unit.kind}\u0000${unit.key}`
}

function spanKey(span: ElementSpan): string {
  return `${span.part}\u0000${span.kind}\u0000${String(span.start)}\u0000${String(span.end)}`
}

function touchedPartsOf(operations: readonly SpliceOperation[]): readonly string[] {
  const parts = new Set<string>()
  for (const operation of operations) {
    if (operation.kind === 'carry') continue
    parts.add(operation.part)
  }
  return [...parts].sort(compareAscii)
}

function buildDispositions(
  input: SplicePlanInput,
  operations: readonly SpliceOperation[],
  removedNames: ReadonlySet<string>,
): readonly PartDisposition[] {
  const dispositions = [...sourceDispositions(input.source, operations, removedNames)]
  for (const entry of input.removedEntries ?? []) {
    dispositions.push({ name: entry.name, status: 'removed' })
  }
  for (const entry of input.addedEntries ?? []) {
    dispositions.push({ name: entry.name, status: 'added' })
  }
  dispositions.sort((left, right) => compareAscii(left.name, right.name))
  return dispositions
}

function sortOperations(operations: readonly SpliceOperation[]): readonly SpliceOperation[] {
  const byPart = new Map<string, SpliceOperation[]>()
  for (const operation of operations) {
    const list = byPart.get(operation.part)
    if (list === undefined) byPart.set(operation.part, [operation])
    else list.push(operation)
  }
  const sorted: SpliceOperation[] = []
  const sortedParts = [...byPart.entries()].sort((left, right) => compareAscii(left[0], right[0]))
  for (const [, list] of sortedParts) {
    const carries = list.filter((operation): operation is CarryOperation => operation.kind === 'carry')
    const byteOperations = list.filter((operation): operation is ByteSpliceOperation => operation.kind !== 'carry')
    if (byteOperations.length > 1) byteOperations.sort(compareByteSpliceOperations)
    sorted.push(...carries, ...byteOperations)
  }
  return sorted
}

function compareByteSpliceOperations(left: ByteSpliceOperation, right: ByteSpliceOperation): number {
  return operationStart(left) - operationStart(right) || operationRank(left) - operationRank(right)
}

function operationStart(operation: ByteSpliceOperation): number {
  return operation.kind === 'insert' || operation.kind === 'append' ? operation.at : operation.start
}

function operationRank(operation: ByteSpliceOperation): number {
  switch (operation.kind) {
    case 'insert': return 0
    case 'append': return 1
    case 'replace': return 2
    case 'delete': return 3
  }
}

function sortEdits(edits: readonly SpliceEdit[]): readonly SpliceEdit[] {
  return [...edits].sort((left, right) =>
    compareAscii(left.part, right.part)
    || left.start - right.start
    || operationRankOfEdit(left) - operationRankOfEdit(right)
    || compareAscii(left.key, right.key))
}

function operationRankOfEdit(edit: SpliceEdit): number {
  if (edit.kind === 'insert') return 0
  if (edit.kind === 'replace') return 1
  return 2
}

function compareAscii(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}

function concatBytes(chunks: readonly Uint8Array[]): Uint8Array {
  let length = 0
  for (const chunk of chunks) length += chunk.byteLength
  const result = new Uint8Array(length)
  let offset = 0
  for (const chunk of chunks) {
    result.set(chunk, offset)
    offset += chunk.byteLength
  }
  return result
}

function refusalPlan(refusal: ConstructIntersection): SplicePlan {
  return { operations: [], touchedParts: [], entries: [], refusal }
}

function countKeys<T>(items: readonly T[], keyOf: (item: T) => string): Map<string, number> {
  const counts = new Map<string, number>()
  for (const item of items) {
    const key = keyOf(item)
    counts.set(key, (counts.get(key) ?? 0) + 1)
  }
  return counts
}

function consumeKey(counts: Map<string, number>, key: string): boolean {
  const count = counts.get(key)
  if (count === undefined) return false
  if (count === 1) counts.delete(key)
  else counts.set(key, count - 1)
  return true
}

function emptyScan(): StructureScan {
  return { properties: [], candidates: [] }
}

function propertyKey(record: PropertyRecord): string {
  return `${record.container}\u0000${elementKey(record.element)}`
}

function elementKey(element: ParsedElement): string {
  return `${element.namespace}\u0000${element.localName}`
}

function mergeRootStyleAttribute(
  sourceBytes: Uint8Array,
  sourceRoot: ParsedElement,
  renderedRoot: ParsedElement | undefined,
  merge: IndependentProperty[],
  rootAttributes: IndependentProperty[],
  coupled: CoupledConstructCandidate[],
): void {
  const cellRoot = sourceRoot.localName === 'c' && isSheetElement(sourceRoot)
  for (const attribute of sourceRoot.attributes) {
    if (attribute.qname === 'xmlns' || attribute.qname.startsWith('xmlns:')) continue
    const reproduced = renderedRoot?.attributes.some(other =>
      other.localName === attribute.localName && other.namespace === attribute.namespace) ?? false
    if (reproduced) continue
    const xml = decodeBytes(sourceBytes.subarray(attribute.start, attribute.end))
    if (cellRoot && attribute.qname === 's') {
      merge.push({
        container: 'c',
        name: attribute.qname,
        xml,
        start: attribute.start,
        end: attribute.end,
      })
      continue
    }
    if (!cellRoot && isIndependentRootMetadata(sourceRoot, attribute)) {
      rootAttributes.push({
        container: 'root',
        name: attribute.qname,
        xml,
        start: attribute.start,
        end: attribute.end,
      })
      continue
    }
    coupled.push({
      name: attribute.qname,
      xml,
      start: attribute.start,
      end: attribute.end,
      detail: `attribute ${attribute.qname} on ${sourceRoot.qname} is not reproduced by the renderer`,
    })
  }
}

function isIndependentRootMetadata(root: ParsedElement, attribute: ParsedAttribute): boolean {
  if (!isWordElement(root)) return false
  if (attribute.localName.startsWith('rsid')) {
    return WORD_ROOT_ATTRIBUTE_NAMESPACES.has(attribute.namespace)
  }
  if (ROOT_EDITOR_ID_NAMES.has(attribute.localName)) {
    return WORD14_ROOT_ATTRIBUTE_NAMESPACES.has(attribute.namespace)
  }
  return false
}

function sliceElement(bytes: Uint8Array, element: ParsedElement): string {
  return decodeBytes(bytes.subarray(element.start, element.end))
}

function renderedParseCandidate(bytes: Uint8Array, detail: string): CoupledConstructCandidate {
  return {
    name: 'xml',
    xml: decodeBytes(bytes),
    start: 0,
    end: bytes.byteLength,
    detail: `rendered unit is not well-formed XML: ${detail}`,
  }
}

function isWordElement(element: ParsedElement): boolean {
  return WORD_NAMESPACES.has(element.namespace) || element.namespace === 'w' || element.namespace === ''
}

function isSheetElement(element: ParsedElement): boolean {
  return SHEET_NAMESPACES.has(element.namespace) || element.namespace === 'x' || element.namespace === ''
}

function propertyContainerOf(element: ParsedElement): PropertyContainer | undefined {
  if (!isWordElement(element) || !PROPERTY_CONTAINERS.has(element.localName)) return undefined
  return element.localName as PropertyContainer
}

function isContentElement(element: ParsedElement): boolean {
  if (WORD_NAMESPACES.has(element.namespace) || element.namespace === 'w') {
    return WORD_CONTENT_ELEMENTS.has(element.localName)
  }
  if (SHEET_NAMESPACES.has(element.namespace) || element.namespace === 'x') {
    return SHEET_CONTENT_ELEMENTS.has(element.localName)
  }
  return WORD_CONTENT_ELEMENTS.has(element.localName) || SHEET_CONTENT_ELEMENTS.has(element.localName)
}

function isSharedFormula(element: ParsedElement): boolean {
  if (!isSheetElement(element) || element.localName !== 'f') return false
  const type = element.attributes.find(attribute => attribute.qname === 't')
  return type !== undefined && SHARED_FORMULA_TYPES.has(type.value)
}

function candidateDetail(element: ParsedElement): string {
  if (REVISION_CHANGE_ELEMENTS.has(element.localName)) {
    return `revision-change metadata ${element.qname} is not reproduced by the renderer`
  }
  if (element.localName === 'ins' || element.localName === 'del'
    || element.localName === 'moveFrom' || element.localName === 'moveTo') {
    return `tracked change ${element.qname} overlaps the edited unit`
  }
  if (element.localName === 'commentRangeStart' || element.localName === 'commentRangeEnd'
    || element.localName === 'commentReference') {
    return `comment range ${element.qname} overlaps the edited unit`
  }
  if (element.localName === 'sdt' || element.localName === 'sdtContent') {
    return `content control ${element.qname} wraps the edited text`
  }
  if (element.localName === 'fldSimple' || element.localName === 'fldChar' || element.localName === 'instrText') {
    return `field markup ${element.qname} owns the edited text`
  }
  if (isSharedFormula(element)) {
    return `shared or array formula ${element.qname} is coupled to the edited cell`
  }
  return `construct ${element.qname} is not reproduced by the renderer`
}

function scanStructure(root: ParsedElement): StructureScan {
  const properties: PropertyRecord[] = []
  const candidates: CandidateRecord[] = []

  const visitProperties = (element: ParsedElement): void => {
    for (const child of element.children) {
      const container = propertyContainerOf(child)
      if (container !== undefined) {
        for (const property of child.children) {
          if (REVISION_CHANGE_ELEMENTS.has(property.localName)) {
            candidates.push({ element: property, detail: candidateDetail(property) })
          } else {
            properties.push({ container, element: property })
          }
        }
        continue
      }
      visitProperties(child)
    }
  }

  const visitCandidates = (element: ParsedElement): void => {
    for (const child of element.children) {
      if (propertyContainerOf(child) !== undefined) continue
      if (isSharedFormula(child)) {
        candidates.push({ element: child, detail: candidateDetail(child) })
        continue
      }
      if (isContentElement(child)) continue
      candidates.push({ element: child, detail: candidateDetail(child) })
      visitCandidates(child)
    }
  }

  visitProperties(root)
  visitCandidates(root)
  return { properties, candidates }
}

function parseUnitXml(bytes: Uint8Array): ParseResult {
  const stack: OpenElement[] = []
  let root: ParsedElement | undefined
  let rootQName = ''
  let index = 0

  while (index < bytes.byteLength) {
    const opening = indexOfByte(bytes, 0x3C, index)
    if (opening === -1) break
    if (startsWith(bytes, opening, COMMENT_OPEN)) {
      const close = findSequence(bytes, COMMENT_CLOSE, opening + COMMENT_OPEN.length)
      if (close === -1) return { ok: false, detail: 'unterminated XML comment' }
      index = close + COMMENT_CLOSE.length
      continue
    }
    if (startsWith(bytes, opening, CDATA_OPEN)) {
      const close = findSequence(bytes, CDATA_CLOSE, opening + CDATA_OPEN.length)
      if (close === -1) return { ok: false, detail: 'unterminated CDATA section' }
      index = close + CDATA_CLOSE.length
      continue
    }
    if (bytes[opening + 1] === 0x3F) {
      const close = findSequence(bytes, PROCESSING_CLOSE, opening + 2)
      if (close === -1) return { ok: false, detail: 'unterminated processing instruction' }
      index = close + PROCESSING_CLOSE.length
      continue
    }
    if (bytes[opening + 1] === 0x21) {
      const close = declarationEnd(bytes, opening + 2)
      if (close === -1) return { ok: false, detail: 'unterminated markup declaration' }
      index = close
      continue
    }
    const tagEnd = findTagEnd(bytes, opening + 1)
    if (tagEnd === -1) return { ok: false, detail: 'unterminated element tag' }
    if (bytes[opening + 1] === 0x2F) {
      const name = readName(bytes, opening + 2, tagEnd)
      if (name === undefined) return { ok: false, detail: 'malformed end tag' }
      const frame = stack.pop()
      if (frame === undefined) return { ok: false, detail: `end tag </${name.name}> has no matching start tag` }
      if (frame.element.qname !== name.name) {
        return { ok: false, detail: `end tag </${name.name}> does not match <${frame.element.qname}>` }
      }
      frame.element.end = tagEnd + 1
      index = tagEnd + 1
      continue
    }
    const parsed = parseStartTag(bytes, opening, tagEnd)
    if (parsed === undefined) return { ok: false, detail: 'malformed start tag' }
    const bindings = new Map<string, string>(stack.at(-1)?.bindings)
    for (const attribute of parsed.attributes) {
      if (attribute.qname === 'xmlns') bindings.set('', attribute.value)
      else if (attribute.qname.startsWith('xmlns:')) bindings.set(attribute.qname.slice('xmlns:'.length), attribute.value)
    }
    const name = splitQName(parsed.qname)
    if (name === undefined) return { ok: false, detail: `malformed element name ${parsed.qname}` }
    const namespace = namespaceToken(name.prefix, bindings)
    const element: ParsedElement = {
      qname: parsed.qname,
      namespace,
      localName: name.localName,
      start: opening,
      end: tagEnd + 1,
      attributes: parsed.attributes.map((attribute) => {
        const attributeName = splitQName(attribute.qname)
        const attributeNamespace = attributeName === undefined || attributeName.prefix === ''
          ? ''
          : namespaceToken(attributeName.prefix, bindings)
        return {
          qname: attribute.qname,
          namespace: attributeNamespace,
          localName: attributeName?.localName ?? attribute.qname,
          value: attribute.value,
          start: attribute.start,
          end: attribute.end,
        }
      }),
      children: [],
    }
    if (stack.length === 0) {
      if (root !== undefined) return { ok: false, detail: 'unit bytes contain more than one root element' }
      root = element
      rootQName = element.qname
    } else {
      stack[stack.length - 1]?.element.children.push(element)
    }
    if (!parsed.selfClosing) stack.push({ element, bindings })
    index = tagEnd + 1
  }

  if (stack.length > 0) {
    return { ok: false, detail: `element <${rootQName}> is not closed` }
  }
  return { ok: true, root }
}

function namespaceToken(prefix: string, bindings: ReadonlyMap<string, string>): string {
  if (prefix === 'xml') return XML_NAMESPACE
  const bound = bindings.get(prefix)
  if (bound !== undefined) return bound
  return prefix === '' ? '' : prefix
}
