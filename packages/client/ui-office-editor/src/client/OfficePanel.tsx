/**
 * The office tab's body: the edited file, its save state, and the editor surface.
 *
 * The file's metadata arrives through the standard `useResource` hook; the
 * document content and its saves travel through the injected `officeDocument`
 * Remote. The editor runtime stays behind the adapter seam and mounts only
 * after a successful read, so loading, read failures, save refusals, external
 * changes, and the dirty close prompt are all visible without loading Univer.
 * The body runs the Session's one durable-set restore pass from the Sidebar's
 * public tab inventory and shows the converted-copy and missing-document
 * notices; the pane's own header lists the open tabs.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { InjectFace, PropsLocale, PropsRuntime } from '@forwork/client-ui-slots'
import type { PreservedConstructKind } from '@forwork/office-exchange/types'
import type {} from '@forwork/client-resources/client'
import type {} from '@forwork/api-workspace-files/client'
import { fileAddressFor } from '@forwork/util-workspace-path'
import type { OfficeEditorHandle } from './adapter.ts'
import type { OfficeCommandTarget } from './command-bridge.ts'
import type { OfficePanelInjected } from './face.ts'
import type { OfficeDocumentFailureCode } from './useOfficeDocument.ts'
import { useOfficeDocument } from './useOfficeDocument.ts'
import { useDocumentRestore } from './restore.ts'
import { openTargetOf, officeDocumentIdentity, projectedDocumentIdAt, workspaceRootOf } from './target.ts'
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

/** Locale key for each preserved construct class. */
const PRESERVED_KEYS: Record<PreservedConstructKind, OfficeEditorKey> = {
  macro: 'preserved.macro',
  chart: 'preserved.chart',
  'pivot-table': 'preserved.pivotTable',
  'excel-table': 'preserved.excelTable',
  'defined-names': 'preserved.definedNames',
  'calc-chain': 'preserved.calcChain',
  comment: 'preserved.comment',
  'embedded-object': 'preserved.embeddedObject',
  drawing: 'preserved.drawing',
  shape: 'preserved.shape',
  theme: 'preserved.theme',
  'font-table': 'preserved.fontTable',
  'web-settings': 'preserved.webSettings',
  'custom-xml': 'preserved.customXml',
  'external-link': 'preserved.externalLink',
  field: 'preserved.field',
  bookmark: 'preserved.bookmark',
  'content-control': 'preserved.contentControl',
  'document-properties': 'preserved.documentProperties',
}

/**
 * The office type's body, registered under `sidebar.right.pane.tab` as `office`.
 * @param props - composed slot props.
 * @returns the file name, the open-document list, save and warning state, and the mounted editor surface.
 */
export function OfficePanel({
  useTabInfo, useProjection, useResource, useSidebarTabs, useRestore, useLegacyNotice, useDocumentLocations, mountEditor,
  attachCommandBridge, markTabDirty, dismissCloseWarning, registerDraftFlush, markMissing, rememberLegacySource,
  dismissLegacyNotice, dismissRestoreNotice, bindDocument, documentIdAt, trackDocument, forgetDocument, observeFile,
  useOfficeTab, remote, restore, t,
}: OfficePanelProps): ReactNode {
  const { tab } = useTabInfo()
  const addressTarget = useMemo(() => openTargetOf(tab.contentId), [tab.contentId])
  const projection = useProjection('document')
  const locations = useDocumentLocations(value => value)
  const bound = documentIdAt(tab.contentId)
  const documentId = useMemo(
    () => bound
      ?? projectedDocumentIdAt(projection, addressTarget.path)
      ?? officeDocumentIdentity(addressTarget.sessionId, addressTarget.path),
    [bound, projection, addressTarget],
  )
  const livePath = locations[documentId] ?? projection?.documents[documentId]?.path ?? addressTarget.path
  const target = useMemo(
    () => ({ sessionId: addressTarget.sessionId, path: livePath }),
    [addressTarget.sessionId, livePath],
  )
  const resourceAddress = useMemo(
    () => livePath === addressTarget.path
      ? tab.contentId
      : fileAddressFor(addressTarget.sessionId, undefined, livePath),
    [livePath, addressTarget, tab.contentId],
  )
  const meta = useResource<'file'>(resourceAddress)
  const workspaceRoot = useMemo(
    () => workspaceRootOf(target.path, meta.value?.absolutePath),
    [target.path, meta.value?.absolutePath],
  )
  const changes = useOfficeTab(value => value)
  const notice = useLegacyNotice(value => value[tab.contentId])
  const openTabs = useSidebarTabs(value => value)
  useDocumentRestore({
    sessionId: target.sessionId,
    projection,
    openTabs,
    bindDocument,
    actions: tab.actions,
    restore,
  })
  useEffect(() => {
    bindDocument(tab.contentId, documentId)
  }, [bindDocument, tab.contentId, documentId])
  useEffect(() => {
    trackDocument(target.sessionId, documentId, target.path)
  }, [trackDocument, target.sessionId, documentId, target.path])
  useEffect(() => () => { forgetDocument(documentId) }, [forgetDocument, documentId])
  useEffect(() => {
    if (meta.status === 'failed') {
      if (meta.failure?.code === 'workspace-file/not-found') {
        observeFile(target.sessionId, documentId, target.path, undefined)
      }
      return
    }
    if (meta.status === 'live' && meta.value !== undefined) {
      observeFile(target.sessionId, documentId, target.path, meta.value.version)
    }
  }, [observeFile, meta.status, meta.failure?.code, meta.value?.version, target.sessionId, documentId, target.path])
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
    documentId,
    signal: tab.signal,
    observedVersion: meta.value?.version,
    workspaceRoot,
    openConverted,
    dirty,
    contentVersion: projection?.documents[documentId]?.contentVersion,
  })
  useEffect(() => {
    if (meta.status !== 'failed' || meta.failure?.code !== 'workspace-file/not-found') return
    markMissing(target.sessionId, documentId)
  }, [markMissing, meta.failure, meta.status, target.sessionId, documentId])
  const surfaceRef = useRef<HTMLDivElement | null>(null)
  const handleRef = useRef<OfficeEditorHandle | null>(null)
  const [mountFailed, setMountFailed] = useState(false)
  const docRef = useRef(doc)
  docRef.current = doc
  const commandTarget = useMemo<OfficeCommandTarget>(() => ({
    documentId,
    version: () => docRef.current.currentContentVersion(),
    isOpen: () => handleRef.current !== null,
    isDirty: () => handleRef.current?.isDirty() ?? false,
    apply: (operation) => {
      const handle = handleRef.current
      if (handle === null) {
        return { status: 'unsupported-feature', features: [{ kind: 'other', detail: 'editor is not mounted' }] }
      }
      const outcome = handle.apply(operation)
      if (outcome.status === 'applied') docRef.current.markApplied()
      return outcome
    },
    snapshot: () => handleRef.current?.getSnapshot(),
  }), [documentId])
  useEffect(
    () => attachCommandBridge(target.sessionId, commandTarget),
    [attachCommandBridge, commandTarget, target.sessionId],
  )
  const flushDraft = useCallback((): void => { void docRef.current.flushDraft() }, [])
  useEffect(
    () => registerDraftFlush(tab.id, flushDraft),
    [registerDraftFlush, tab.id, flushDraft],
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
    let started = false
    let handle: OfficeEditorHandle | undefined
    let observer: ResizeObserver | undefined
    // A hidden tab lays out its surface at 0x0, so Univer mounts without a
    // caret; wait for a real box and mount once as soon as the tab is visible.
    const start = (): void => {
      if (started || !live) return
      started = true
      observer?.disconnect()
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
    }
    if (typeof ResizeObserver === 'undefined') {
      start()
    } else {
      observer = new ResizeObserver(() => {
        if (container.clientWidth > 0 && container.clientHeight > 0) start()
      })
      observer.observe(container)
      if (container.clientWidth > 0 && container.clientHeight > 0) start()
    }
    return () => {
      observer?.disconnect()
      live = false
      if (handle !== undefined) handle.dispose()
      if (handleRef.current === handle) handleRef.current = null
      markTabDirty(tab.id, false)
    }
  }, [doc.snapshot, doc.recovered, mountEditor, onChange, onEdit, markTabDirty, tab.id])

  const persist = useCallback((closeAfter: boolean): void => {
    void doc.save(handleRef.current?.getSnapshot()).then((outcome) => {
      if (outcome !== 'saved') return
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
      {doc.preserved.length > 0 && (
        <div className={css.preserved} data-office-preserved>
          <p className={css.preservedLine}>{t('preserved')}</p>
          <ul className={css.features}>
            {doc.preserved.map(kind => (
              <li key={kind} data-office-preserved-feature={kind}>{t(PRESERVED_KEYS[kind])}</li>
            ))}
          </ul>
        </div>
      )}
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
