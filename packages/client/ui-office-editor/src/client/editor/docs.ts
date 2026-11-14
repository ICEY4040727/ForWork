/**
 * The Univer Docs runtime: the core preset, its English pack, and the CSS the
 * preset requires. Loaded only through the adapter's dynamic import.
 */
import { CommandType, LocaleType } from '@univerjs/core'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import UniverDocsCoreEnUS from '@univerjs/preset-docs-core/locales/en-US'
import '@univerjs/preset-docs-core/lib/index.css'
import { createUniver } from '@univerjs/presets'
import type { OfficeRuntime, OfficeRuntimeMount } from '../adapter.ts'
import { nativeHistory } from '../undo.ts'
import { documentData, jsonValueOf, localePack } from '../wire.ts'

/**
 * Mount a word-processing document into one container.
 * @param input - the container and the wire snapshot to load.
 * @returns the live runtime behind the adapter.
 * @throws when the snapshot is not a document.
 */
export function mount(input: OfficeRuntimeMount): Promise<OfficeRuntime> {
  if (input.snapshot.kind !== 'document') {
    throw new Error('ui-office-editor: a workbook snapshot cannot mount in the docs runtime')
  }
  const { container, snapshot } = input
  const data = documentData(snapshot.data)
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: localePack(UniverDocsCoreEnUS as Record<string, unknown>),
    },
    presets: [UniverDocsCorePreset({ container })],
  })
  const unit = univerAPI.createDocument(data)
  const history = nativeHistory(univerAPI)
  const listeners = new Set<() => void>()
  const notify = (): void => {
    for (const listener of [...listeners]) listener()
  }
  const events = [
    univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
      if (event.type === CommandType.MUTATION) notify()
    }),
    univerAPI.addEvent(univerAPI.Event.Undo, notify),
    univerAPI.addEvent(univerAPI.Event.Redo, notify),
  ]
  return Promise.resolve({
    getSnapshot: () => ({ kind: 'document', data: jsonValueOf(unit.save()) }),
    undo: () => { history.undo() },
    redo: () => { history.redo() },
    onMutation(listener) {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    dispose() {
      for (const event of events) event.dispose()
      listeners.clear()
      univer.dispose()
      container.replaceChildren()
    },
  })
}
