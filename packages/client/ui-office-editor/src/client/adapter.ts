/**
 * The editor adapter seam: the one boundary the Univer presets cross.
 *
 * The component layer never imports an editor runtime. An adapter mounts a
 * surface into a container through an injected runtime factory, owns the live
 * unit, and reports dirty changes; the default factory dynamically imports the
 * sheets or docs module only when a document actually mounts. Component tests
 * replace the factory so jsdom never loads Univer.
 */
import type { OfficeDocumentSnapshot } from '@deepseek-ai/dsh-office-exchange/types'
import type { DocumentEditOperation } from '@deepseek-ai/dsh-session/src/document-events.ts'
import type { OfficeApplyOutcome } from './command-bridge.ts'
import { loadOfficeRuntime } from './editor/load.ts'

/** The unit kind an adapter opens, matching the `officeDocument` snapshot discriminant. */
export type OfficeAdapterKind = 'workbook' | 'document'

/** The Remote wire snapshot of one Office document, as `officeDocument.read` returns it and `save` accepts it. */
export type OfficeSnapshot = OfficeDocumentSnapshot

/** What one mount receives from the owning panel. */
export interface OfficeMountInput {
  /** The opened unit kind. */
  readonly kind: OfficeAdapterKind
  /** The snapshot the surface starts from. */
  readonly snapshot: OfficeSnapshot
  /** Called on the first edit after the unit became clean, to mark the tab dirty. */
  onChange(): void
  /** Called after every user edit, so recovery drafts capture the latest content. */
  onEdit?(): void
}

/** What one runtime mount receives: the target element and the snapshot to load. */
export interface OfficeRuntimeMount {
  /** The element the editor surface belongs to. */
  readonly container: HTMLElement
  /** The opened unit kind. */
  readonly kind: OfficeAdapterKind
  /** The native snapshot the runtime starts from. */
  readonly snapshot: OfficeSnapshot
}

/** One live editor runtime behind the adapter. */
export interface OfficeRuntime {
  /** The edited unit as a Remote wire snapshot. */
  getSnapshot(): OfficeSnapshot
  /**
   * Apply one structured operation to the live unit.
   * @param operation - ordered steps from the Host command.
   * @returns whether every step applied, or the features that stop it.
   */
  apply(operation: DocumentEditOperation): OfficeApplyOutcome
  /** Undo through the editor's native history. */
  undo(): void
  /** Redo through the editor's native history. */
  redo(): void
  /**
   * Observe commands that change the persisted model.
   * @param listener - called after a mutation, undo, or redo.
   * @returns removes the listener.
   */
  onMutation(listener: () => void): () => void
  /** Tear the editor instance down and release its listeners. */
  dispose(): void
}

/** Loads and constructs one editor runtime. The default dynamically imports Univer. */
export type OfficeRuntimeFactory = (mount: OfficeRuntimeMount) => Promise<OfficeRuntime>

/** One mounted editor surface. */
export interface OfficeEditorHandle {
  /** Remove the surface and release every listener; repeated calls are no-ops. */
  dispose(): void
  /** The edited unit snapshot, in the same wire form `officeDocument.save` accepts. */
  getSnapshot(): OfficeSnapshot
  /**
   * Apply one structured operation to the live unit.
   * @param operation - ordered steps from the Host command.
   * @returns whether every step applied, or the features that stop it.
   */
  apply(operation: DocumentEditOperation): OfficeApplyOutcome
  /** Undo the last edit through the editor's own stack. */
  undo(): void
  /** Redo the last undone edit through the editor's own stack. */
  redo(): void
  /** Whether the live unit has edits the file does not hold yet. */
  isDirty(): boolean
  /** Mark the live unit as saved without touching its content. */
  clearDirty(): void
}

/** The editor runtime behind the seam. */
export interface OfficeAdapter {
  /**
   * Mount one editor surface.
   * @param container - the element the surface belongs to.
   * @param input - opened kind, starting snapshot, and change callback.
   * @returns the mounted handle.
   */
  mount(container: HTMLElement, input: OfficeMountInput): Promise<OfficeEditorHandle>
}

/**
 * Create an adapter over one runtime factory.
 * @param factory - runtime loader; the default dynamically imports the Univer preset modules.
 * @returns the adapter the tab body mounts.
 */
export function createAdapter(factory: OfficeRuntimeFactory = loadOfficeRuntime): OfficeAdapter {
  return {
    async mount(container, input) {
      const runtime = await factory({ container, kind: input.kind, snapshot: input.snapshot })
      let dirty = false
      let applying = false
      let disposed = false
      const release = runtime.onMutation(() => {
        if (applying) return
        input.onEdit?.()
        if (dirty) return
        dirty = true
        input.onChange()
      })
      return {
        dispose() {
          if (disposed) return
          disposed = true
          release()
          runtime.dispose()
          container.replaceChildren()
        },
        getSnapshot: () => runtime.getSnapshot(),
        apply(operation) {
          applying = true
          try {
            return runtime.apply(operation)
          } finally {
            applying = false
          }
        },
        undo: () => { runtime.undo() },
        redo: () => { runtime.redo() },
        isDirty: () => dirty,
        clearDirty: () => { dirty = false },
      }
    },
  }
}

/** The shipped adapter: the default factory loads the preset modules lazily. */
export const officeAdapter: OfficeAdapter = createAdapter()

/**
 * A controllable adapter for tests, which never load an editor runtime.
 *
 * `mounts` records the inputs in order; `setDirty` simulates one edit on the
 * most recently mounted handle and notifies both edit and dirty listeners.
 */
export interface MockOfficeAdapter extends OfficeAdapter {
  /** Mount inputs in mount order. */
  readonly mounts: readonly OfficeMountInput[]
  /** Operations the mounted handles applied, in call order. */
  readonly applied: readonly DocumentEditOperation[]
  /**
   * Set the most recently mounted handle's dirty state and notify its owner.
   * @param dirty - the new state; a call before any mount does nothing.
   */
  setDirty(dirty: boolean): void
}

/**
 * Create a test adapter that records mounts and exposes dirty control.
 * @returns the mock adapter.
 */
export function createMockAdapter(): MockOfficeAdapter {
  const mounts: OfficeMountInput[] = []
  const applied: DocumentEditOperation[] = []
  const states: { input: OfficeMountInput; dirty: boolean }[] = []
  return {
    mounts,
    applied,
    setDirty(dirty) {
      const state = states.at(-1)
      if (state === undefined) return
      state.dirty = dirty
      state.input.onEdit?.()
      state.input.onChange()
    },
    mount(container, input) {
      const state = { input, dirty: false }
      states.push(state)
      mounts.push(input)
      container.dataset.officeMount = input.kind
      return Promise.resolve({
        dispose() {
          delete container.dataset.officeMount
        },
        getSnapshot: () => input.snapshot,
        apply(operation) {
          applied.push(operation)
          return { status: 'applied' }
        },
        undo() {},
        redo() {},
        isDirty: () => state.dirty,
        clearDirty() {
          state.dirty = false
        },
      })
    },
  }
}
