/**
 * The change surface behind the dirty close veto and the external-change rule,
 * plus the tab-address translation. No DOM and no editor runtime.
 */
import { describe, expect, it, vi } from 'vitest'
import { brandString } from '@forwork/brand'
import type { TabId } from '@forwork/client-ui-dockkit'
import type { SessionId } from '@forwork/session/types'
import type { DocumentId } from '@forwork/session/src/document-events.ts'
import type { DocumentProjectionView } from '@forwork/session-document/client'
import { createOfficeTabChanges, externalChangeVerdict } from '../src/client/changes.ts'
import { openTargetOf, projectedDocumentIdAt } from '../src/client/target.ts'

const TAB = brandString<TabId>('tab-1')
const OTHER = brandString<TabId>('tab-2')
const SESSION = brandString<SessionId>('s-1')
const OTHER_SESSION = brandString<SessionId>('s-2')

describe('createOfficeTabChanges', () => {
  it('tracks dirty and warned state per tab', () => {
    const changes = createOfficeTabChanges()
    changes.markDirty(TAB, true)
    expect(changes.isDirty(TAB)).toBe(true)
    expect(changes.source.getSnapshot()[TAB]).toEqual({ dirty: true, warned: false })
    changes.warnClose(TAB)
    expect(changes.source.getSnapshot()[TAB]).toEqual({ dirty: true, warned: true })
    changes.dismissWarning(TAB)
    expect(changes.source.getSnapshot()[TAB]).toEqual({ dirty: true, warned: false })
    changes.markDirty(TAB, false)
    expect(changes.source.getSnapshot()[TAB]).toEqual({ dirty: false, warned: false })
    expect(changes.isDirty(TAB)).toBe(false)
  })

  it('ignores warnings on a clean tab and keeps other tabs untouched', () => {
    const changes = createOfficeTabChanges()
    changes.warnClose(TAB)
    expect(changes.source.getSnapshot()[TAB]).toBeUndefined()
    changes.dismissWarning(TAB)
    changes.forget(TAB)
    changes.markDirty(TAB, true)
    changes.markDirty(OTHER, true)
    changes.warnClose(TAB)
    changes.forget(TAB)
    expect(changes.source.getSnapshot()[TAB]).toBeUndefined()
    expect(changes.source.getSnapshot()[OTHER]).toEqual({ dirty: true, warned: false })
    changes.forget(OTHER)
    expect(changes.source.getSnapshot()).toEqual({})
  })

  it('keeps a warned tab dirty and clears the warning with the dirty state', () => {
    const changes = createOfficeTabChanges()
    changes.markDirty(TAB, true)
    changes.warnClose(TAB)
    changes.markDirty(TAB, true)
    expect(changes.source.getSnapshot()[TAB]).toEqual({ dirty: true, warned: true })
    changes.markDirty(TAB, false)
    expect(changes.source.getSnapshot()[TAB]).toEqual({ dirty: false, warned: false })
  })
})

describe('externalChangeVerdict', () => {
  it('ignores missing versions and a metadata frame older than the read', () => {
    expect(externalChangeVerdict({ observedVersion: undefined, version: 'v1', observedAtRead: 'v1', dirty: false }))
      .toBe('none')
    expect(externalChangeVerdict({ observedVersion: 'v2', version: undefined, observedAtRead: undefined, dirty: false }))
      .toBe('none')
    expect(externalChangeVerdict({ observedVersion: 'v1', version: 'v2', observedAtRead: 'v1', dirty: false }))
      .toBe('none')
    expect(externalChangeVerdict({ observedVersion: 'v2', version: 'v2', observedAtRead: 'v1', dirty: false }))
      .toBe('none')
  })

  it('reloads a clean unit and prompts a dirty one', () => {
    expect(externalChangeVerdict({ observedVersion: 'v2', version: 'v1', observedAtRead: 'v1', dirty: false }))
      .toBe('reload')
    expect(externalChangeVerdict({ observedVersion: 'v2', version: 'v1', observedAtRead: 'v1', dirty: true }))
      .toBe('prompt')
  })
})

describe('document move tracking', () => {
  const DOCUMENT = `${SESSION}:work/budget.xlsx`
  const OLD = 'work/budget.xlsx'
  const NEW = 'work/renamed.xlsx'

  function tracked(rebind = vi.fn()) {
    const changes = createOfficeTabChanges(rebind)
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    changes.observeFile(SESSION, DOCUMENT, OLD, 'v1')
    return { changes, rebind }
  }

  it('rebinds the document when its version reappears after its path disappears', () => {
    const { changes, rebind } = tracked()

    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    expect(rebind).not.toHaveBeenCalled()

    changes.observeFile(SESSION, DOCUMENT, NEW, 'v1')
    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION, DOCUMENT, NEW)
    changes.observeFile(SESSION, DOCUMENT, NEW, 'v1')
    expect(rebind).toHaveBeenCalledOnce()
  })

  it('resolves a disappearance against a path observed before it', () => {
    const { changes, rebind } = tracked()
    const other = `${SESSION}:work/copy.xlsx`
    changes.trackDocument(SESSION, other, NEW)
    changes.observeFile(SESSION, other, NEW, 'v1')

    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION, DOCUMENT, NEW)
  })

  it('resolves a missing document when another open tab observes the same version', () => {
    const { changes, rebind } = tracked()
    const other = `${SESSION}:work/copy.xlsx`
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.trackDocument(SESSION, other, NEW)
    changes.observeFile(SESSION, other, NEW, 'v1')
    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION, DOCUMENT, NEW)
  })

  it('ignores a different version and another Session', () => {
    const { changes, rebind } = tracked()
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.observeFile(SESSION, DOCUMENT, NEW, 'v2')
    changes.trackDocument(OTHER_SESSION, `${OTHER_SESSION}:work/other.xlsx`, 'work/other.xlsx')
    changes.observeFile(OTHER_SESSION, `${OTHER_SESSION}:work/other.xlsx`, 'work/other.xlsx', 'v1')
    expect(rebind).not.toHaveBeenCalled()
  })

  it('keeps a document that reappears at its own path', () => {
    const { changes, rebind } = tracked()
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.observeFile(SESSION, DOCUMENT, OLD, 'v1')
    expect(rebind).not.toHaveBeenCalled()
  })

  it('stops following a document whose tab is gone', () => {
    const { changes, rebind } = tracked()
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.forgetDocument(DOCUMENT)
    changes.observeFile(SESSION, DOCUMENT, NEW, 'v1')
    expect(rebind).not.toHaveBeenCalled()
  })

  it('ignores an observation for a document it does not follow', () => {
    const rebind = vi.fn()
    const changes = createOfficeTabChanges(rebind)
    changes.observeFile(SESSION, 'unknown', OLD, undefined)
    changes.observeFile(SESSION, 'unknown', NEW, 'v1')
    expect(rebind).not.toHaveBeenCalled()
  })

  it('tracks an adopted path change and normalizes separators', () => {
    const rebind = vi.fn()
    const changes = createOfficeTabChanges(rebind)
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    changes.trackDocument(SESSION, DOCUMENT, 'work\\next.xlsx')
    changes.observeFile(SESSION, DOCUMENT, 'work/next.xlsx', 'v1')
    changes.observeFile(SESSION, DOCUMENT, 'work\\next.xlsx', undefined)
    changes.observeFile(SESSION, DOCUMENT, 'work/final.xlsx', 'v1')
    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION, DOCUMENT, 'work/final.xlsx')
  })

  it('replaces a document identity tracked under another Session', () => {
    const rebind = vi.fn()
    const changes = createOfficeTabChanges(rebind)
    changes.trackDocument(OTHER_SESSION, DOCUMENT, OLD)
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    changes.observeFile(SESSION, DOCUMENT, OLD, 'v1')
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.observeFile(SESSION, DOCUMENT, NEW, 'v1')
    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION, DOCUMENT, NEW)
  })

  it('ignores an absence for a path the document no longer addresses', () => {
    const { changes, rebind } = tracked()
    changes.observeFile(SESSION, DOCUMENT, NEW, undefined)
    expect(rebind).not.toHaveBeenCalled()
  })

  it('marks an unobserved version missing without a rebind candidate', () => {
    const rebind = vi.fn()
    const changes = createOfficeTabChanges(rebind)
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    expect(rebind).not.toHaveBeenCalled()
  })

  it('does not match an observed path holding another version', () => {
    const { changes, rebind } = tracked()
    changes.observeFile(SESSION, DOCUMENT, NEW, 'v2')
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    expect(rebind).not.toHaveBeenCalled()
  })

  it('resolves a missing document when another tab reports a third path', () => {
    const { changes, rebind } = tracked()
    const other = `${SESSION}:work/copy.xlsx`
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.trackDocument(SESSION, other, NEW)
    changes.observeFile(SESSION, other, NEW, 'v2')
    changes.observeFile(SESSION, other, 'work/third.xlsx', 'v1')
    expect(rebind).toHaveBeenCalledExactlyOnceWith(SESSION, DOCUMENT, 'work/third.xlsx')
  })

  it('still follows documents when no rebind sink is installed', () => {
    const changes = createOfficeTabChanges()
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    changes.observeFile(SESSION, DOCUMENT, OLD, 'v1')
    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    expect(() => { changes.observeFile(SESSION, DOCUMENT, NEW, 'v1') }).not.toThrow()
  })

  it('publishes the moved path under the unchanged document identity', () => {
    const { changes } = tracked()
    expect(changes.locations.getSnapshot()).toEqual({ [DOCUMENT]: OLD })

    changes.observeFile(SESSION, DOCUMENT, OLD, undefined)
    changes.observeFile(SESSION, DOCUMENT, NEW, 'v1')

    expect(changes.locations.getSnapshot()).toEqual({ [DOCUMENT]: NEW })
    changes.forgetDocument(DOCUMENT)
    changes.forgetDocument(DOCUMENT)
    expect(changes.locations.getSnapshot()).toEqual({})
  })

  it('publishes a location when a document is first followed or re-addressed', () => {
    const changes = createOfficeTabChanges()
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    expect(changes.locations.getSnapshot()).toEqual({ [DOCUMENT]: OLD })

    const settled = changes.locations.getSnapshot()
    changes.trackDocument(SESSION, DOCUMENT, OLD)
    expect(changes.locations.getSnapshot()).toBe(settled)

    changes.trackDocument(SESSION, DOCUMENT, 'work/next.xlsx')
    expect(changes.locations.getSnapshot()).toEqual({ [DOCUMENT]: 'work/next.xlsx' })
  })
})

describe('document identity bindings', () => {
  const ADDRESS = 'dsh-resource://file/session/s-1/work/budget.xlsx'
  const DOCUMENT = `${SESSION}:work/budget.xlsx`

  it('binds an address to the projection identity and releases it on close', () => {
    const changes = createOfficeTabChanges()
    expect(changes.documentIdAt(ADDRESS)).toBeUndefined()

    changes.bindDocument(ADDRESS, DOCUMENT)
    expect(changes.documentIdAt(ADDRESS)).toBe(DOCUMENT)

    changes.releaseDocument(ADDRESS, `${SESSION}:other.xlsx`)
    expect(changes.documentIdAt(ADDRESS)).toBe(DOCUMENT)

    changes.releaseDocument(ADDRESS, DOCUMENT)
    expect(changes.documentIdAt(ADDRESS)).toBeUndefined()
  })
})

describe('projectedDocumentIdAt', () => {
  const PRESENT = brandString<DocumentId>(`${SESSION}:work/budget.xlsx`)
  const view: DocumentProjectionView = {
    order: [brandString<DocumentId>(`${SESSION}:dangling.xlsx`), PRESENT],
    documents: {
      [PRESENT]: {
        documentId: PRESENT,
        path: 'work/budget.xlsx',
        format: 'xlsx',
        contentVersion: 1,
        status: 'open',
        unsupported: [],
      },
    },
  }

  it('finds the projected identity at a locator and skips a dangling order entry', () => {
    expect(projectedDocumentIdAt(view, 'work/budget.xlsx')).toBe(PRESENT)
    expect(projectedDocumentIdAt(view, 'work/other.xlsx')).toBeUndefined()
    expect(projectedDocumentIdAt(undefined, 'work/budget.xlsx')).toBeUndefined()
  })
})

describe('openTargetOf', () => {
  it('reads the Session and path from a session file address', () => {
    expect(openTargetOf('dsh-resource://file/session/s-1/work/budget.xlsx'))
      .toEqual({ sessionId: 's-1', path: 'work/budget.xlsx' })
  })

  it('refuses an address that names no Session', () => {
    expect(() => openTargetOf('dsh-resource://file/absolute/work/budget.xlsx'))
      .toThrow('not a Session-scoped file address')
  })
})
