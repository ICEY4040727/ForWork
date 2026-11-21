/**
 * The `data.officeExchange.source` sidecar and its validation rules.
 *
 * The sidecar binds an editable snapshot to the source package it was read
 * from. It carries either one anchor per editable unit or one summary per
 * sheet, never both, and every span must lie inside a real part with a
 * non-overlapping byte range. Validation happens at the JSON boundary: a
 * malformed sidecar throws {@link SourceStateError}, which the Remote maps to
 * `conversion-failed`, so an invalid anchor set can never reach the splice
 * planner.
 * @module @forwork/office-exchange/source/source-state
 */

/** Formats an in-place save accepts. */
export type OfficeSourceFormat = 'docx' | 'docm' | 'xlsx' | 'xlsm'

/** Unit kinds a per-unit anchor can name. */
export type UnitAnchorKind = 'paragraph' | 'tableCell' | 'sheetCell' | 'contentControl'

/** Which anchor representation the sidecar carries. */
export type SourceAnchorMode = 'per-unit' | 'sheet-summary'

/** One editable unit in one source part. */
export interface UnitAnchor {
  /** Unit kind the span covers. */
  readonly kind: UnitAnchorKind
  /** Decoded part name, for example `word/document.xml`. */
  readonly part: string
  /** Model identity: `paragraphId`, `${tableId}/${row}/${cell}`, or `${sheetId}!${row}:${column}`. */
  readonly key: string
  /** Inclusive byte offset of the unit in the uncompressed part. */
  readonly start: number
  /** Exclusive byte offset of the unit. */
  readonly end: number
  /** Lowercase SHA-256 hex of the canonical unit projection at read. */
  readonly digest: string
}

/** Read-time summary used instead of per-cell anchors for a large sheet. */
export interface SheetSummary {
  /** Model sheet id. */
  readonly sheetId: string
  /** Resolved worksheet part. */
  readonly part: string
  /** Lowercase SHA-256 hex of the canonical sheet projection at read. */
  readonly digest: string
  /** Non-empty cells at read. */
  readonly cellCount: number
}

interface OfficeSourceStateHeader {
  readonly schemaVersion: 1
  readonly format: OfficeSourceFormat
  readonly packageDigest: string
  readonly byteLength: number
  readonly sourceParts: readonly string[]
}

/** Sidecar stored at `data.officeExchange.source` inside the editable data object. */
export type OfficeSourceState =
  | (OfficeSourceStateHeader & {
    readonly anchorMode: 'per-unit'
    readonly anchors: readonly UnitAnchor[]
  })
  | (OfficeSourceStateHeader & {
    readonly anchorMode: 'sheet-summary'
    readonly sheets: readonly SheetSummary[]
  })

/** Stable reasons a splice plan cannot be completed. */
export type PreservationFailureKind =
  | 'ambiguous-span'
  | 'unsupported-coupling'
  | 'digest-mismatch'
  | 'anchor-budget-exceeded'
  | 'unsupported-container'
  | 'encoding-unsupported'

/** A typed reason a plan cannot be completed, carrying its location when one exists. */
export interface PreservationFailure {
  /** Stable classification consumers branch on. */
  readonly kind: PreservationFailureKind
  /** Actionable description of the construct and its location. */
  readonly detail: string
  /** Part the failure concerns, when it concerns one part. */
  readonly part?: string
  /** Unit key the failure concerns, when it concerns one unit. */
  readonly key?: string
}

/** Stable parser failure codes. */
export type SourceStateErrorCode =
  | 'invalid-state'
  | 'unknown-part'
  | 'invalid-span'
  | 'overlapping-spans'
  | 'duplicate-key'
  | 'invalid-digest'
  | 'anchor-budget-exceeded'

/** Thrown when a sidecar violates a structural, span, identity, digest, or budget rule. */
export class SourceStateError extends Error {
  /**
   * @param code - Stable classification for callers that branch on the rule that failed.
   * @param message - Actionable description naming the offending field or span.
   */
  constructor(
    readonly code: SourceStateErrorCode,
    message: string,
  ) {
    super(message)
    this.name = 'SourceStateError'
  }
}

/**
 * Source facts and limits the sidecar cannot state about itself.
 *
 * `partByteLengths` is both the part-name set and the per-part bound for anchor
 * spans. Budgets default to the values in the data model; a caller that
 * resolves configuration passes its validated values explicitly.
 */
export interface SourceStateContext {
  /** Uncompressed byte length of every part in the source package, keyed by decoded part name. */
  readonly partByteLengths: ReadonlyMap<string, number>
  /** Largest accepted number of unit records. Defaults to {@link DEFAULT_MAX_SOURCE_ANCHORS}. */
  readonly maxSourceAnchors?: number
  /** Largest accepted UTF-8 byte length of the unit-record payload. Defaults to {@link DEFAULT_MAX_SOURCE_ANCHOR_BYTES}. */
  readonly maxSourceAnchorBytes?: number
}

/** Default cap on the number of per-unit anchors. */
export const DEFAULT_MAX_SOURCE_ANCHORS = 100_000

/** Default cap on the serialized anchor payload, in bytes. */
export const DEFAULT_MAX_SOURCE_ANCHOR_BYTES = 8 * 1024 * 1024

const SOURCE_FORMATS: ReadonlySet<string> = new Set(['docx', 'docm', 'xlsx', 'xlsm'])
const ANCHOR_MODES: ReadonlySet<string> = new Set(['per-unit', 'sheet-summary'])
const ANCHOR_KINDS: ReadonlySet<string> = new Set(['paragraph', 'tableCell', 'sheetCell', 'contentControl'])
const DIGEST_PATTERN = /^[0-9a-f]{64}$/u

/**
 * Validate a decoded sidecar and rebuild it without unknown fields.
 * @param value - Untrusted value from the editable snapshot's `data.officeExchange.source`.
 * @param context - Source part lengths and the validated anchor count and byte budgets.
 * @returns A sidecar whose anchors or sheets satisfy every data-model rule.
 * @throws SourceStateError When a field, span, identity, digest, or budget rule fails.
 * @throws RangeError When a caller-supplied budget is not a positive safe integer.
 */
export function parseSourceState(value: unknown, context: SourceStateContext): OfficeSourceState {
  const maxAnchors = context.maxSourceAnchors ?? DEFAULT_MAX_SOURCE_ANCHORS
  const maxBytes = context.maxSourceAnchorBytes ?? DEFAULT_MAX_SOURCE_ANCHOR_BYTES
  assertBudget('maxSourceAnchors', maxAnchors)
  assertBudget('maxSourceAnchorBytes', maxBytes)
  const record = requireRecord(value, 'source state')
  if (record.schemaVersion !== 1) invalid('invalid-state', 'schemaVersion must be 1')
  const format = requireFormat(record.format)
  const packageDigest = requireDigest(record.packageDigest, 'packageDigest')
  const byteLength = requireNonNegativeInteger(record.byteLength, 'byteLength')
  const sourceParts = parseSourceParts(record.sourceParts, context.partByteLengths)
  const anchorMode = requireAnchorMode(record.anchorMode)
  if (anchorMode === 'per-unit') {
    if (record.sheets !== undefined) invalid('invalid-state', 'sheets must be absent in per-unit mode')
    if (record.anchors === undefined) invalid('invalid-state', 'anchors must be present in per-unit mode')
    const anchors = parseAnchors(record.anchors, context.partByteLengths)
    assertPayloadBudget(anchors, maxAnchors, maxBytes)
    return { schemaVersion: 1, format, packageDigest, byteLength, sourceParts, anchorMode, anchors }
  }
  if (record.anchors !== undefined) invalid('invalid-state', 'anchors must be absent in sheet-summary mode')
  if (record.sheets === undefined) invalid('invalid-state', 'sheets must be present in sheet-summary mode')
  const sheets = parseSheets(record.sheets, context.partByteLengths)
  assertPayloadBudget(sheets, maxAnchors, maxBytes)
  return { schemaVersion: 1, format, packageDigest, byteLength, sourceParts, anchorMode, sheets }
}

function invalid(code: SourceStateErrorCode, detail: string): never {
  throw new SourceStateError(code, detail)
}

function assertBudget(field: string, value: number): void {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(`${field} must be a positive integer, got ${String(value)}`)
  }
}

function assertPayloadBudget(units: readonly unknown[], maxAnchors: number, maxBytes: number): void {
  if (units.length > maxAnchors) {
    throw new SourceStateError(
      'anchor-budget-exceeded',
      `unit record count ${String(units.length)} exceeds maxSourceAnchors ${String(maxAnchors)}`,
    )
  }
  const payloadBytes = Buffer.byteLength(JSON.stringify(units), 'utf8')
  if (payloadBytes > maxBytes) {
    throw new SourceStateError(
      'anchor-budget-exceeded',
      `unit payload is ${String(payloadBytes)} bytes, exceeding maxSourceAnchorBytes ${String(maxBytes)}`,
    )
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function requireRecord(value: unknown, path: string): Record<string, unknown> {
  if (!isRecord(value)) invalid('invalid-state', `${path} must be an object`)
  return value
}

function requireArray(value: unknown, path: string): readonly unknown[] {
  if (!Array.isArray(value)) invalid('invalid-state', `${path} must be an array`)
  const items: unknown[] = value
  return items
}

function requireString(value: unknown, path: string): string {
  if (typeof value !== 'string') invalid('invalid-state', `${path} must be a string`)
  return value
}

function requireNonEmptyString(value: unknown, path: string): string {
  const text = requireString(value, path)
  if (text.length === 0) invalid('invalid-state', `${path} must not be empty`)
  return text
}

function requireInteger(value: unknown, path: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value)) {
    invalid('invalid-state', `${path} must be a safe integer`)
  }
  return value
}

function requireNonNegativeInteger(value: unknown, path: string): number {
  const number = requireInteger(value, path)
  if (number < 0) invalid('invalid-state', `${path} must be non-negative`)
  return number
}

function requireFormat(value: unknown): OfficeSourceFormat {
  if (typeof value !== 'string' || !SOURCE_FORMATS.has(value)) {
    invalid('invalid-state', 'format must be docx, docm, xlsx, or xlsm')
  }
  return value as OfficeSourceFormat
}

function requireAnchorMode(value: unknown): SourceAnchorMode {
  if (typeof value !== 'string' || !ANCHOR_MODES.has(value)) {
    invalid('invalid-state', 'anchorMode must be per-unit or sheet-summary')
  }
  return value as SourceAnchorMode
}

function requireAnchorKind(value: unknown, path: string): UnitAnchorKind {
  if (typeof value !== 'string' || !ANCHOR_KINDS.has(value)) {
    invalid('invalid-state', `${path} must be paragraph, tableCell, sheetCell, or contentControl`)
  }
  return value as UnitAnchorKind
}

function requireDigest(value: unknown, path: string): string {
  const digest = requireString(value, path)
  if (!DIGEST_PATTERN.test(digest)) {
    invalid('invalid-digest', `${path} must be 64 lowercase hex characters`)
  }
  return digest
}

function parseSourceParts(value: unknown, partByteLengths: ReadonlyMap<string, number>): readonly string[] {
  const items = requireArray(value, 'sourceParts')
  const seen = new Set<string>()
  const parts: string[] = []
  for (const item of items) {
    const part = requireNonEmptyString(item, 'sourceParts[]')
    if (!partByteLengths.has(part)) invalid('unknown-part', `sourceParts contains unknown part ${part}`)
    if (seen.has(part)) invalid('invalid-state', `sourceParts repeats ${part}`)
    seen.add(part)
    parts.push(part)
  }
  return parts
}

function parseAnchors(value: unknown, partByteLengths: ReadonlyMap<string, number>): readonly UnitAnchor[] {
  const items = requireArray(value, 'anchors')
  const anchors = items.map((item, index) => parseAnchor(item, index, partByteLengths))
  const identities = new Set<string>()
  for (const anchor of anchors) {
    const identity = JSON.stringify([anchor.part, anchor.kind, anchor.key])
    if (identities.has(identity)) {
      invalid('duplicate-key', `anchor key ${anchor.key} repeats for ${anchor.kind} in ${anchor.part}`)
    }
    identities.add(identity)
  }
  assertNoOverlaps(anchors)
  return anchors
}

function parseAnchor(value: unknown, index: number, partByteLengths: ReadonlyMap<string, number>): UnitAnchor {
  const path = `anchors[${String(index)}]`
  const record = requireRecord(value, path)
  const kind = requireAnchorKind(record.kind, `${path}.kind`)
  const part = requireString(record.part, `${path}.part`)
  const partByteLength = partByteLengths.get(part)
  if (partByteLength === undefined) invalid('unknown-part', `${path}.part is not a source part: ${part}`)
  const key = requireNonEmptyString(record.key, `${path}.key`)
  const start = requireInteger(record.start, `${path}.start`)
  const end = requireInteger(record.end, `${path}.end`)
  if (start < 0) invalid('invalid-span', `${path}.start must be non-negative`)
  if (start >= end) invalid('invalid-span', `${path}.end must be greater than start`)
  if (end > partByteLength) invalid('invalid-span', `${path}.end exceeds ${part} length ${String(partByteLength)}`)
  const digest = requireDigest(record.digest, `${path}.digest`)
  return { kind, part, key, start, end, digest }
}

function assertNoOverlaps(anchors: readonly UnitAnchor[]): void {
  const byPart = new Map<string, UnitAnchor[]>()
  for (const anchor of anchors) {
    const list = byPart.get(anchor.part)
    if (list === undefined) byPart.set(anchor.part, [anchor])
    else list.push(anchor)
  }
  for (const [part, list] of byPart) {
    const sorted = [...list].sort((left, right) => left.start - right.start || left.end - right.end)
    let previous: UnitAnchor | undefined
    for (const anchor of sorted) {
      if (previous !== undefined && anchor.start < previous.end) {
        invalid('overlapping-spans', `anchors overlap in ${part} at byte ${String(anchor.start)}`)
      }
      previous = anchor
    }
  }
}

function parseSheets(value: unknown, partByteLengths: ReadonlyMap<string, number>): readonly SheetSummary[] {
  const items = requireArray(value, 'sheets')
  const seen = new Set<string>()
  return items.map((item, index) => {
    const path = `sheets[${String(index)}]`
    const record = requireRecord(item, path)
    const sheetId = requireNonEmptyString(record.sheetId, `${path}.sheetId`)
    if (seen.has(sheetId)) invalid('invalid-state', `${path}.sheetId repeats ${sheetId}`)
    seen.add(sheetId)
    const part = requireString(record.part, `${path}.part`)
    if (!partByteLengths.has(part)) invalid('unknown-part', `${path}.part is not a source part: ${part}`)
    const digest = requireDigest(record.digest, `${path}.digest`)
    const cellCount = requireNonNegativeInteger(record.cellCount, `${path}.cellCount`)
    return { sheetId, part, digest, cellCount }
  })
}
