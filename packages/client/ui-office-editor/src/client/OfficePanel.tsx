/**
 * The office tab's body: the edited file, its save state, and the editor surface.
 *
 * The file's metadata arrives through the standard `useResource` hook; the
 * document content and its saves travel through the injected `officeDocument`
 * Remote. The editor runtime stays behind the adapter seam and mounts only
 * after a successful read, so loading, read failures, save refusals, external
 * changes, and the dirty close prompt are all visible without loading Univer.
 * The body also lists the Session's other open Office documents from the
 * Sidebar's public tab inventory, runs the Session's one durable-set restore
 * pass, and shows the converted-copy and missing-document notices.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { InjectFace, PropsLocale, PropsRuntime } from '@forwork/client-ui-slots'
import type { UnsupportedFeatureKind } from '@forwork/office-exchange/types'
import type {} from '@forwork/client-resources/client'
import type {} from '@forwork/api-workspace-files/client'
import { pathPartsOf } from '@forwork/util-workspace-path'
import type { OfficeEditorHandle } from './adapter.ts'
import type { OfficeCommandTarget } from './command-bridge.ts'
import type { OfficePanelInjected } from './face.ts'
import { DocumentPanel } from './DocumentPanel.tsx'
import type { OfficeDocumentEntry } from './DocumentPanel.tsx'
import type { OfficeDocumentFailureCode } from './useOfficeDocument.ts'
import { useOfficeDocument } from './useOfficeDocument.ts'
import { useDocumentRestore } from './restore.ts'
import { OFFICE_KIND, basenameOf } from './definition.ts'
import { openTargetOf, officeDocumentIdentity, workspaceRootOf } from './target.ts'
import type { OfficeEditorKey } from './locales.ts'
import css from './OfficePanel.module.css'

/** The body's composed props: the tab seat, the injected face, and this namespace's copy. */
export type OfficePanelProps =
  & PropsRuntime<'sidebar.right.pane.tab'>
  & InjectFace<OfficePanelInjected>
  & PropsLocale<'officeEditor'>

/** The missing-path list shared by every panel with nothing to report. */
const NO_MISSING: readonly string[] = []

/** Locale key for each save-refusal class. */
const ERROR_KEYS: Record<OfficeDocumentFailureCode, OfficeEditorKey> = {
  conflict: 'error.conflict',
  'unsupported-format': 'error.unsupportedFormat',
  'unsupported-feature': 'error.unsupportedFeature',
  'out-of-scope': 'error.outOfScope',
  'bound-exceeded': 'error.boundExceeded',
  'conversion-failed': 'error.conversionFailed',
  cancelled: 'error.cancelled',
  unknown: 'error.unknown',
}

/** Locale key for each unsupported feature class. */
const FEATURE_KEYS: Record<UnsupportedFeatureKind, OfficeEditorKey> = {
  macro: 'feature.macro',
  'tracked-changes': 'feature.trackedChanges',
  comment: 'feature.comment',
  chart: 'feature.chart',
  drawing: 'feature.drawing',
  shape: 'feature.shape',
  'header-footer': 'feature.headerFooter',
  'formula-recalc': 'feature.formulaRecalc',
  'embedded-object': 'feature.embeddedObject',
  other: 'feature.other',
}

/**
 * The office type's body, registered under `sidebar.right.pane.tab` as `office`.
 * @param props - composed slot props.
 * @returns the file name, the open-document list, save and warning state, and the mounted editor surface.
 */
export function OfficePanel({
  useTabInfo, useProjection, useResource, useSidebarTabs, useRestore, useLegacyNotice, mountEditor, attachCommandBridge,
  markTabDirty, dismissCloseWarning, markMissing, rememberLegacySource, dismissLegacyNotice, dismissRestoreNotice,
  useOfficeTab, remote, restore, t,
}: OfficePanelProps): ReactNode {
  const { tab } = useTabInfo()
  const meta = useResource<'file'>(tab.contentId)
  const target = useMemo(() => openTargetOf(tab.contentId), [tab.contentId])
  const workspaceRoot = useMemo(
    () => workspaceRootOf(target.path, meta.value?.absolutePath),
    [target.path, meta.value?.absolutePath],
  )
  const changes = useOfficeTab(value => value)
  const notice = useLegacyNotice(value => value[tab.contentId])
  const openTabs = useSidebarTabs(value => value)
  const projection = useProjection('document')
  useDocumentRestore({
    sessionId: target.sessionId,
    projection,
    openTabs,
    actions: tab.actions,
    restore,
  })
  const restored = useRestore(value => value)
  const missing = restored.missing[target.sessionId] ?? NO_MISSING
  const dirty = changes[tab.id]?.dirty === true
  const warned = changes[tab.id]?.warned === true
  const openConverted = useCallback((address: string, source: string): void => {
    rememberLegacySource(address, source)
    tab.actions.openResource(address, { replaceTab: true })
  }, [rememberLegacySource, tab.actions])
  const doc = useOfficeDocument({
    remote,
    sessionId: target.sessionId,
    path: target.path,
    signal: tab.signal,
    observedVersion: meta.value?.version,
    workspaceRoot,
    openConverted,
    dirty,
  })
  useEffect(() => {
    if (meta.status !== 'failed' || meta.failure?.code !== 'workspace-file/not-found') return
    markMissing(target.sessionId, officeDocumentIdentity(target.sessionId, target.path))
  }, [markMissing, meta.failure, meta.status, target.sessionId, target.path])
  const documents = useMemo<readonly OfficeDocumentEntry[]>(() => openTabs
    .filter(entry => entry.sessionId === target.sessionId && entry.kind === OFFICE_KIND)
    .map(entry => ({
      tabId: entry.tabId,
      contentId: entry.contentId,
      title: basenameOf(entry.contentId),
      active: entry.tabId === tab.id,
    })), [openTabs, target.sessionId, tab.id])
  const activateDocument = useCallback((entry: OfficeDocumentEntry): void => {
    tab.actions.openResource(entry.contentId)
  }, [tab.actions])
  const surfaceRef = useRef<HTMLDivElement | null>(null)
  const handleRef = useRef<OfficeEditorHandle | null>(null)
  const versionRef = useRef(0)
  const [mountFailed, setMountFailed] = useState(false)
  const name = meta.value === undefined ? tab.title : pathPartsOf(meta.value.absolutePath).name
  const docRef = useRef(doc)
  docRef.current = doc
  const commandTarget = useMemo<OfficeCommandTarget>(() => ({
    documentId: officeDocumentIdentity(target.sessionId, target.path),
    version: () => versionRef.current,
    isOpen: () => handleRef.current !== null,
    isDirty: () => handleRef.current?.isDirty() ?? false,
    unsupported: () => docRef.current.unsupported,
    apply: (operation) => {
      const handle = handleRef.current
      if (handle === null) {
        return { status: 'unsupported-feature', features: [{ kind: 'other', detail: 'editor is not mounted' }] }
      }
      const outcome = handle.apply(operation)
      if (outcome.status === 'applied') versionRef.current += 1
      return outcome
    },
    snapshot: () => handleRef.current?.getSnapshot(),
  }), [target.sessionId, target.path])
  useEffect(
    () => attachCommandBridge(target.sessionId, commandTarget),
    [attachCommandBridge, commandTarget, target.sessionId],
  )
  const onChange = useCallback((): void => { markTabDirty(tab.id, true) }, [markTabDirty, tab.id])
  const onEdit = useCallback((): void => {
    const live = handleRef.current?.getSnapshot()
    if (live !== undefined) docRef.current.scheduleDraft(live)
  }, [])

  useEffect(() => {
    const container = surfaceRef.current
    const snapshot = doc.snapshot
    if (container === null || snapshot === undefined) return
    let live = true
    let handle: OfficeEditorHandle | undefined
    versionRef.current = 0
    setMountFailed(false)
    void mountEditor(container, { kind: snapshot.kind, snapshot, onChange, onEdit }).then((mounted) => {
      if (!live) {
        mounted.dispose()
        return
      }
      handle = mounted
      handleRef.current = mounted
      if (doc.recovered) markTabDirty(tab.id, true)
    }, () => {
      if (live) setMountFailed(true)
    })
    return () => {
      live = false
      if (handle !== undefined) handle.dispose()
      if (handleRef.current === handle) handleRef.current = null
      markTabDirty(tab.id, false)
    }
  }, [doc.snapshot, doc.recovered, mountEditor, onChange, onEdit, markTabDirty, tab.id])

  const persist = useCallback((closeAfter: boolean): void => {
    void doc.save(handleRef.current?.getSnapshot()).then((outcome) => {
      if (outcome !== 'saved') return
      versionRef.current += 1
      handleRef.current?.clearDirty()
      markTabDirty(tab.id, false)
      if (closeAfter) tab.actions.close()
    })
  }, [doc, markTabDirty, tab.actions, tab.id])

  const discardAndClose = useCallback((): void => {
    markTabDirty(tab.id, false)
    dismissCloseWarning(tab.id)
    doc.discardDraft()
    tab.actions.close()
  }, [dismissCloseWarning, doc, markTabDirty, tab.actions, tab.id])

  const statusLine = doc.status === 'loading'
    ? <p className={css.statusLine} data-office-loading>{t('loading')}</p>
    : doc.status === 'converting'
      ? <p className={css.statusLine} data-office-converting>{t('converting')}</p>
      : doc.status === 'error' && doc.error !== undefined
        ? <p className={css.statusLine} data-office-error={doc.error.code}>{t(ERROR_KEYS[doc.error.code])}</p>
        : mountFailed
          ? <p className={css.statusLine} data-office-mount-failed>{t('editorUnavailable')}</p>
          : null
  return (
    <section className={css.panel} data-office-editor-panel>
      <header className={css.header}>
        <span className={css.fileName} data-office-file-name>{name}</span>
        <span className={css.actions}>
          {dirty && <span className={css.dirty} data-office-dirty>{t('unsaved')}</span>}
          <button
            type="button"
            className={css.save}
            data-office-save
            disabled={!doc.canSave || mountFailed}
            onClick={() => { persist(false) }}
          >
            {doc.saving ? t('saving') : t('save')}
          </button>
        </span>
      </header>
      {notice !== undefined && !notice.dismissed && (
        <p className={css.banner} data-office-legacy-notice>
          <span>{t('legacyNotice', { source: notice.source })}</span>
          <button
            type="button"
            className={css.bannerAction}
            data-office-dismiss-legacy
            onClick={() => { dismissLegacyNotice(tab.contentId) }}
          >
            {t('dismiss')}
          </button>
        </p>
      )}
      {missing.length > 0 && (
        <p className={css.banner} data-office-restore-missing>
          <span>{t('restoreMissing', { paths: missing.join(', ') })}</span>
          <button
            type="button"
            className={css.bannerAction}
            data-office-dismiss-restore
            onClick={() => { dismissRestoreNotice(target.sessionId) }}
          >
            {t('dismiss')}
          </button>
        </p>
      )}
      <DocumentPanel documents={documents} onActivate={activateDocument} t={t} />
      {doc.recovery !== undefined && (
        <p className={css.banner} data-office-draft-recovery={doc.recovery.decision}>
          <span>{doc.recovery.decision === 'stale' ? t('draftStale') : t('draftFound')}</span>
          <button
            type="button"
            className={css.bannerAction}
            data-office-recover-draft
            onClick={() => { doc.acceptRecovery() }}
          >
            {doc.recovery.decision === 'stale' ? t('recoverAnyway') : t('recover')}
          </button>
          <button
            type="button"
            className={css.bannerAction}
            data-office-discard-draft
            onClick={() => { doc.discardRecovery() }}
          >
            {t('discardDraft')}
          </button>
        </p>
      )}
      {doc.draftError !== undefined && (
        <p className={css.banner} data-office-draft-error>
          <span>{t('draftError')}</span>
          <button
            type="button"
            className={css.bannerAction}
            data-office-dismiss-draft-error
            onClick={() => { doc.dismissDraftError() }}
          >
            {t('dismiss')}
          </button>
        </p>
      )}
      {doc.unsupported.length > 0 && (
        <div className={css.refusal} data-office-unsupported>
          <p className={css.refusalLine}>{t('unsupported')}</p>
          <ul className={css.features}>
            {doc.unsupported.map(feature => (
              <li key={`${feature.kind}:${feature.detail ?? ''}`} data-office-unsupported-feature={feature.kind}>
                <span>{t(FEATURE_KEYS[feature.kind])}</span>
                {feature.detail !== undefined && <span className={css.featureDetail}>{feature.detail}</span>}
              </li>
            ))}
          </ul>
        </div>
      )}
      {doc.externalChange && (
        <p className={css.banner} data-office-external-change>
          <span>{t('externalChange')}</span>
          <button type="button" className={css.bannerAction} onClick={() => { doc.reload() }}>{t('reload')}</button>
          <button type="button" className={css.bannerAction} onClick={() => { doc.dismissExternalChange() }}>{t('keepEditing')}</button>
        </p>
      )}
      {doc.status === 'conflict' && doc.error !== undefined && (
        <p className={css.banner} data-office-conflict>
          <span>{t(ERROR_KEYS[doc.error.code])}</span>
          <button type="button" className={css.bannerAction} onClick={() => { doc.reload() }}>{t('reload')}</button>
        </p>
      )}
      {doc.status !== 'conflict' && doc.status !== 'error' && doc.error !== undefined && (
        <p className={css.banner} data-office-save-error>
          <span>{t(ERROR_KEYS[doc.error.code])}</span>
        </p>
      )}
      {warned && dirty && (
        <p className={css.banner} data-office-close-warning>
          <span>{t('closeWarning')}</span>
          <button
            type="button"
            className={css.bannerAction}
            data-office-save-close
            disabled={!doc.canSave || mountFailed}
            onClick={() => { persist(true) }}
          >
            {t('saveAndClose')}
          </button>
          <button type="button" className={css.bannerAction} data-office-discard-close onClick={discardAndClose}>
            {t('discardAndClose')}
          </button>
          <button
            type="button"
            className={css.bannerAction}
            data-office-dismiss-close
            onClick={() => { dismissCloseWarning(tab.id) }}
          >
            {t('cancel')}
          </button>
        </p>
      )}
      {statusLine}
      <div className={css.surface} data-office-surface>
        <div className={css.mount} ref={surfaceRef} data-office-mount />
      </div>
    </section>
  )
}
