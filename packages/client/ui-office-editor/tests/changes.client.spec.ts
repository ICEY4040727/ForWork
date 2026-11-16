/**
 * The change surface behind the dirty close veto and the external-change rule,
 * plus the tab-address translation. No DOM and no editor runtime.
 */
import { describe, expect, it } from 'vitest'
import { brandString } from '@forwork/brand'
import type { TabId } from '@forwork/client-ui-dockkit'
import { createOfficeTabChanges, externalChangeVerdict } from '../src/client/changes.ts'
import { openTargetOf } from '../src/client/target.ts'

const TAB = brandString<TabId>('tab-1')
const OTHER = brandString<TabId>('tab-2')

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
