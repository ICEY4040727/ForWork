/**
 * Office Exchange capability seam (`ctx.officeExchange`) and its local provider.
 * The seam maps Office files to and from unit snapshots and reports features the
 * editable model cannot preserve; the editor and agent tools consume it without
 * importing a concrete conversion library. The package also hosts the
 * `officeDocument` Remote surface over that seam.
 * @module @forwork/office-exchange
 */

export { assertWithinRoots } from './authority.ts'
export type { WriteRoots } from './authority.ts'
export { OfficeCommandBroker } from './commands.ts'
export type { DocumentCommandRequest } from './commands.ts'
export { detectPreserved } from './detect.ts'
export { OfficeExchange } from './service.ts'
export { OfficeExchangeError } from './errors.ts'
export type { OfficeExchangeErrorCode } from './errors.ts'
export { formatOf, isEditableFormat, isLegacyFormat } from './format.ts'
export { Config, LocalOfficeExchange } from './provider/index.ts'
export { OfficeDocumentRemote, RemoteConfig } from './remote.ts'
export { default } from './provider/index.ts'
export type { SourcePackage } from './package/types.ts'
export type { ExchangeReadResult, OfficeWriteResult, UnitSnapshot } from './snapshots.ts'
export type {
  CoupledConstructKind,
  CoupledConstructMarker,
  DocumentSaveAttribution,
  EditableOfficeFormat,
  LegacyOfficeFormat,
  OfficeCommandOperation,
  OfficeCommandOutcome,
  OfficeCommandResult,
  OfficeDocumentDraft,
  OfficeDocumentDraftWriteResult,
  OfficeDocumentReadResult,
  OfficeDocumentSaveResult,
  OfficeDocumentSnapshot,
  OfficeFormat,
  PreservedConstructKind,
  PreservedConstructs,
  SaveDisposition,
  SaveDispositionGroup,
  UnitSnapshotKind,
  UnsupportedFeature,
  UnsupportedFeatureKind,
} from './types.ts'
