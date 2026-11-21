/** Stable tab siblings in one horizontal Grid, including viewport-positioned floats. */
import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { MutableRefObject, ReactNode } from 'react'
import clsx from 'clsx'
import type { DockIntents } from '../contract/adapter.ts'
import type { LayoutState, PaneId, PaneNode, TabId, TabRecord } from '../contract/types.ts'
import { findTabPane, floatRect, getNode, getPane, getTab } from '../engine/tree.ts'
import type { SizePreview } from './PaneTree.tsx'
import type { PaneCallbacks } from './render.ts'
import { FloatHeader, floatStyle, useFloatGestures } from './FloatLayer.tsx'
import { TabBody } from './TabBody.tsx'
import { PaneDropHints, TabPanel, TabStrip } from './TabPanel.tsx'
import css from './dockkit.module.css'

/** Host policy for lazily retaining a tab body without reparenting it. */
export interface TabRetention {
  /** Whether a visited body survives hiding; omitted means visibility-mounted. */
  readonly keepMounted?: (tab: TabRecord) => boolean
  /** Whether this layout's Session is on screen; defaults to true. */
  readonly active?: boolean
}

interface LayoutProps extends TabRetention {
  readonly state: LayoutState
  readonly callbacks: PaneCallbacks
  readonly preview: SizePreview | undefined
  readonly intents: DockIntents
}

interface TabHostProps extends LayoutProps {
  readonly tab: TabRecord
  readonly pane: PaneNode
  readonly column: number
  readonly floats: ReturnType<typeof useFloatGestures>
  readonly focusRequest: MutableRefObject<{ readonly tabId: TabId; readonly origin: Element | null } | undefined>
  /** Chip-box offsets by pane, read and recorded by each selected tab's strip. */
  readonly offsets: Map<PaneId, number>
}

/**
 * Focus a tab's remembered element and give a retained editor a usable caret.
 * A hidden contenteditable can keep focusability while losing its selection,
 * so restoring focus alone leaves typed characters without an input event.
 * @param element - focus target inside the tab's section.
 */
function focusRestored(element: HTMLElement): void {
  element.focus({ preventScroll: true })
  if (!element.isContentEditable) return
  const selection = window.getSelection()
  // Only a document with a browsing context exposes a selection.
  if (selection === null) return
  if (selection.rangeCount > 0 && element.contains(selection.anchorNode)) return
  const range = document.createRange()
  range.selectNodeContents(element)
  range.collapse(false)
  selection.removeAllRanges()
  selection.addRange(range)
}

/** A tab's ancestors stay identical across selection, pane moves and floating. */
function TabHost({ state, callbacks, intents, tab, pane, column, floats, focusRequest, offsets, keepMounted,
  active = true }: TabHostProps): ReactNode {
  const floating = pane.host === 'float'
  const selected = floating || pane.activeTabId === tab.id
  const visible = active && selected && (floating || state.expanded)
  const retained = keepMounted?.(tab) ?? false
  const [visited, setVisited] = useState(visible)
  if (visible && !visited) setVisited(true)
  const host = useRef<HTMLElement | null>(null)
  const body = useRef<HTMLDivElement | null>(null)
  const lastFocused = useRef<HTMLElement | null>(null)
  useLayoutEffect(() => {
    // Both refs target unconditional descendants attached before these effects.
    const section = host.current as HTMLElement
    section.inert = !visible
    const focused = document.activeElement
    if (!visible) {
      // A retained body keeps its mounted editor; its hidden input must be
      // blurred now and focused again when the tab returns.
      if (focused instanceof HTMLElement && section.contains(focused)) {
        lastFocused.current = focused
        focused.blur()
      }
      return
    }
    const restore = (): void => {
      const target = lastFocused.current
      if (target?.isConnected === true && section.contains(target)) focusRestored(target)
      lastFocused.current = null
    }
    const request = focusRequest.current
    if (request?.tabId === tab.id) {
      focusRequest.current = undefined
      // Preserve deliberate focus taken by another control or the newly shown body.
      if (focused !== null && focused !== document.body && focused !== document.documentElement && focused !== request.origin) return
      const strip = section.querySelector('[data-dockkit-strip]')
      const chip = [...strip?.querySelectorAll<HTMLElement>('[data-dockkit-tab]') ?? []]
        .find(element => element.dataset.dockkitTab === tab.id)
      if (chip !== undefined) {
        focusRestored(chip)
        lastFocused.current = null
      } else {
        // A custom pane header can replace the strip; activating its tab then
        // returns to the editor the tab held before it was hidden.
        restore()
      }
      return
    }
    if (focused === null || focused === document.body || focused === document.documentElement) restore()
  }, [visible, tab.id, focusRequest])
  // A click on another tab's chip moves focus before this tab hides, so the
  // body remembers its own focus target while it is still on screen.
  useEffect(() => {
    const element = body.current as HTMLDivElement
    const remember = (event: FocusEvent): void => {
      if (event.target instanceof HTMLElement) lastFocused.current = event.target
    }
    element.addEventListener('focusin', remember)
    return () => { element.removeEventListener('focusin', remember) }
  }, [])
  // Electron emits a non-bubbling focus event on the webview element.
  useEffect(() => {
    const element = body.current as HTMLDivElement
    const focus = (): void => {
      if (!visible) return
      if (floating) floats.raise(pane.id)
      else if (state.activePaneId !== pane.id) callbacks.onFocusPane(pane.id)
    }
    element.addEventListener('focus', focus, true)
    return () => { element.removeEventListener('focus', focus, true) }
  }, [visible, floating, floats, pane.id, state.activePaneId, callbacks])
  const lifted = floating && floats.preview?.paneId === pane.id ? floats.preview.rect : undefined
  const rect = floating ? floats.rectOf(pane.id, floatRect(pane)) : undefined
  const paneHeader = selected && !floating
    ? callbacks.renderPaneHeader?.(tab, pane, pane.tabs.map(id => getTab(state, id)))
    : undefined
  const depth = lifted === undefined ? state.floats.indexOf(pane.id) + 1 : state.floats.length + 1
  return (
    <div className={clsx(css.tabCell, floating && css.floatingCell)} hidden={!selected}
      data-dockkit-host={floating ? 'float' : 'dock'} data-dockkit-column={floating ? undefined : column}
      style={{ gridColumn: floating ? 1 : column * 2 + 1, gridRow: 1, order: floating ? depth : 0 }}>
      <section ref={host} tabIndex={-1} className={clsx(css.tabHost, floating ? css.float : css.pane)}
        aria-hidden={!visible || undefined}
        data-dockkit-content={tab.id}
        data-dockkit-pane={!floating && selected ? pane.id : undefined}
        data-dockkit-pane-active={!floating && state.activePaneId === pane.id || undefined}
        data-dockkit-float={floating ? pane.id : undefined}
        data-dockkit-float-active={floating && state.activePaneId === pane.id || undefined}
        data-dockkit-column={!floating ? column : undefined}
        style={rect === undefined ? undefined : floatStyle(rect, lifted !== undefined)}
        onPointerDown={() => { if (floating) floats.raise(pane.id) }}
        onClick={() => { if (!floating && state.activePaneId !== pane.id) callbacks.onFocusPane(pane.id) }}>
        <div className={css.tabHostHeader}>
          {selected && (floating
            ? <FloatHeader paneId={pane.id} tab={tab} labels={callbacks.labels} intents={intents}
              renderTabTitle={callbacks.renderTabTitle} canCloseTab={callbacks.canCloseTab} drag={floats.drag} />
            : paneHeader === undefined
              ? <TabStrip state={state} pane={pane} callbacks={callbacks} offsets={offsets} />
              : paneHeader)}
        </div>
        <div ref={body} className={clsx(css.tabHostBody, floating ? css.floatBody : css.paneBody)}>
          {visited && (retained || (active && selected)) ? <TabBody tab={tab} renderTab={callbacks.renderTab} /> : null}
          {!floating && <PaneDropHints pane={pane} callbacks={callbacks} />}
        </div>
        {floating && <div className={css.floatResize} data-dockkit-float-resize={pane.id}
          onPointerDown={(event) => { floats.drag('resize', pane.id, event) }} />}
      </section>
    </div>
  )
}

/**
 * One or two horizontal panes; CSS resolves content sizes from track fractions.
 * @param props - committed layout, gesture preview and body retention policy.
 * @returns stable tab containers, including floating tabs.
 * @throws if the docked tree is not one pane or two horizontal panes.
 */
export function TabLayout(props: LayoutProps): ReactNode {
  const { state, callbacks, preview } = props
  const focusRequest = useRef<{ readonly tabId: TabId; readonly origin: Element | null }>()
  // Each pane's strip remounts under the newly selected tab's host.
  const [offsets] = useState(() => new Map<PaneId, number>())
  const tabCallbacks: PaneCallbacks = {
    ...callbacks,
    onFocusTab: (tabId) => {
      // Only changing the selected body replaces its strip; pane-only focus does not.
      focusRequest.current = findTabPane(state, tabId).activeTabId === tabId
        ? undefined : { tabId, origin: document.activeElement }
      callbacks.onFocusTab(tabId)
    },
  }
  const root = getNode(state, state.rootId)
  if (root.kind === 'split' && (root.axis !== 'row' || root.children.length !== 2)) {
    throw new Error('DockLayout requires one pane or two horizontally split panes')
  }
  const panes = root.kind === 'pane' ? [root] : root.children.map(id => getPane(state, id))
  const sizes = root.kind === 'split' ? preview?.splitId === root.id ? preview.sizes : root.sizes : [1]
  const floats = useFloatGestures(state, props.intents)
  const columns = sizes.map(size => `minmax(0, ${size}fr)`).join(' 0px ')
  return (
    <div className={css.tabLayout} data-dockkit-split={root.kind === 'split' ? root.id : undefined}
      style={{ gridTemplateColumns: columns }}>
      {/* Logical order changes must not make React move a connected webview's ancestor. */}
      {Object.values(state.tabs).sort((a, b) => a.id.localeCompare(b.id)).map((tab) => {
        const pane = findTabPane(state, tab.id)
        return <TabHost key={tab.id} {...props} callbacks={tabCallbacks} tab={tab} pane={pane}
          column={panes.findIndex(candidate => candidate.id === pane.id)} floats={floats} focusRequest={focusRequest} offsets={offsets} />
      })}
      {panes.filter(pane => pane.tabs.length === 0).map(pane => (
        <div key={pane.id} className={css.emptyTabHost} data-dockkit-empty style={{ gridColumn: panes.indexOf(pane) * 2 + 1, gridRow: 1 }}>
          <TabPanel state={state} pane={pane} callbacks={callbacks} />
        </div>
      ))}
      {root.kind === 'split' && <div className={clsx(css.divider, css.tabLayoutDivider)}
        data-dockkit-divider={`${root.id}:0`} style={{ gridColumn: 2, gridRow: 1 }}
        onPointerDown={(event) => { callbacks.onDividerPressed(root.id, 0, event) }} />}
    </div>
  )
}
