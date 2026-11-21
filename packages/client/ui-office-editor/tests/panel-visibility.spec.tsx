// @vitest-environment jsdom
/**
 * The editor surface mounts only once its container has a laid-out box.
 *
 * A hidden tab measures 0x0, and Univer mounted there never computes a caret.
 * The panel observes the mount container and starts the editor on the first
 * non-empty box; without a `ResizeObserver` it keeps the immediate mount.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { Mock } from 'vitest'
import { act, cleanup, waitFor } from '@testing-library/react'
import type { RenderResult } from '@testing-library/react'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import type { RemoteResult } from '@forwork/typert-protocol'
import type { OfficeDocumentReadResult, OfficeDocumentSnapshot } from '@forwork/office-exchange/types'
import { createMockAdapter } from '../src/client/adapter.ts'
import type { OfficeMountInput } from '../src/client/adapter.ts'
import type { OfficePanelInjected } from '../src/client/face.ts'
import type { OfficeDocumentRemote } from '../src/client/useOfficeDocument.ts'
import { renderPanel } from './panel.tsx'

afterEach(() => {
  cleanup()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

const snapshot: OfficeDocumentSnapshot = { kind: 'workbook', data: { id: 'book' } }

function success<T>(value: T): RemoteResult<T> {
  return { ok: true, value }
}

function fakeRemote(): OfficeDocumentRemote {
  return {
    read: vi.fn(async (): Promise<RemoteResult<OfficeDocumentReadResult>> => success({
      format: 'xlsx',
      snapshot,
      unsupported: [],
      version: 'v1',
    })),
    convertLegacy: vi.fn(),
    save: vi.fn(),
    writeDraft: vi.fn(),
    readDraft: vi.fn(async (): Promise<RemoteResult<null>> => success(null)),
    clearDraft: vi.fn(async (): Promise<RemoteResult<void>> => success(undefined)),
  }
}

let observers: ResizeObserverStub[]

class ResizeObserverStub {
  readonly observed: Element[] = []
  disconnected = false

  constructor(private readonly callback: ResizeObserverCallback) {
    observers.push(this)
  }

  observe(target: Element): void {
    this.observed.push(target)
  }

  unobserve(): void {}

  disconnect(): void {
    this.disconnected = true
  }

  fire(): void {
    this.callback([], this)
  }
}

function onlyObserver(): ResizeObserverStub {
  const observer = observers[0]
  if (observer === undefined) throw new Error('the panel created no ResizeObserver')
  return observer
}

function mountContainer(view: RenderResult): HTMLElement {
  const container = view.container.querySelector<HTMLElement>('[data-office-mount]')
  if (container === null) throw new Error('the panel rendered no mount container')
  return container
}

function sizeMount(view: RenderResult, width: number, height: number): void {
  const container = mountContainer(view)
  Object.defineProperty(container, 'clientWidth', { configurable: true, value: width })
  Object.defineProperty(container, 'clientHeight', { configurable: true, value: height })
}

function mountSpy(): { adapter: ReturnType<typeof createMockAdapter>; mount: Mock<OfficePanelInjected['mountEditor']> } {
  const adapter = createMockAdapter()
  const mount = vi.fn<OfficePanelInjected['mountEditor']>(
    (container: HTMLElement, input: OfficeMountInput) => adapter.mount(container, input),
  )
  return { adapter, mount }
}

beforeEach(() => {
  observers = []
  vi.stubGlobal('ResizeObserver', ResizeObserverStub)
})

describe('editor mount visibility', () => {
  it('does not mount while the hidden surface has no box', async () => {
    const { adapter, mount } = mountSpy()
    const view = renderPanel({ remote: fakeRemote(), mount })
    await waitFor(() => { expect(observers).toHaveLength(1) })
    expect(onlyObserver().observed).toEqual([mountContainer(view)])
    expect(mount).not.toHaveBeenCalled()

    act(() => { onlyObserver().fire() })
    expect(mount).not.toHaveBeenCalled()
    sizeMount(view, 640, 0)
    act(() => { onlyObserver().fire() })
    expect(mount).not.toHaveBeenCalled()
    sizeMount(view, 0, 480)
    act(() => { onlyObserver().fire() })
    expect(mount).not.toHaveBeenCalled()
    expect(adapter.mounts).toHaveLength(0)
  })

  it('mounts once when the surface gains a box and ignores later callbacks', async () => {
    const { adapter, mount } = mountSpy()
    const view = renderPanel({ remote: fakeRemote(), mount })
    await waitFor(() => { expect(observers).toHaveLength(1) })
    expect(mount).not.toHaveBeenCalled()

    sizeMount(view, 640, 480)
    act(() => { onlyObserver().fire() })
    await waitFor(() => { expect(adapter.mounts).toHaveLength(1) })
    expect(mount).toHaveBeenCalledOnce()
    expect(onlyObserver().disconnected).toBe(true)

    act(() => { onlyObserver().fire() })
    await act(async () => {})
    expect(mount).toHaveBeenCalledOnce()
    expect(adapter.mounts).toHaveLength(1)
  })

  it('mounts from the first size check when the tab is already visible', async () => {
    const { adapter, mount } = mountSpy()
    vi.spyOn(Element.prototype, 'clientWidth', 'get').mockReturnValue(640)
    vi.spyOn(Element.prototype, 'clientHeight', 'get').mockReturnValue(480)
    renderPanel({ remote: fakeRemote(), mount })
    await waitFor(() => { expect(adapter.mounts).toHaveLength(1) })
    expect(mount).toHaveBeenCalledOnce()
  })

  it('ignores an observer callback after the body unmounts', async () => {
    const { adapter, mount } = mountSpy()
    const view = renderPanel({ remote: fakeRemote(), mount })
    await waitFor(() => { expect(observers).toHaveLength(1) })
    sizeMount(view, 640, 480)
    view.unmount()
    act(() => { onlyObserver().fire() })
    await act(async () => {})
    expect(mount).not.toHaveBeenCalled()
    expect(adapter.mounts).toHaveLength(0)
  })

  it('mounts immediately where the environment has no ResizeObserver', async () => {
    vi.stubGlobal('ResizeObserver', undefined)
    const { adapter, mount } = mountSpy()
    renderPanel({ remote: fakeRemote(), mount })
    await waitFor(() => { expect(adapter.mounts).toHaveLength(1) })
    expect(mount).toHaveBeenCalledOnce()
  })
})

describe('office panel stylesheet bytes', () => {
  it('starts with the panel selector and carries no UTF-8 byte-order mark', () => {
    const bytes = readFileSync(resolve(import.meta.dirname, '../src/client/OfficePanel.module.css'))
    expect([...bytes.subarray(0, 3)]).not.toEqual([0xef, 0xbb, 0xbf])
    expect(bytes.subarray(0, 6).toString('utf8')).toBe('.panel')
  })
})
