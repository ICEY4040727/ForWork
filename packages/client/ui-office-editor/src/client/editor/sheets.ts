/**
 * The Univer Sheets runtime: the core preset, its English pack, and the CSS
 * the preset requires. Loaded only through the adapter's dynamic import.
 */
import { CommandType, LocaleType } from '@univerjs/core'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver } from '@univerjs/presets'
import type { OfficeRuntime, OfficeRuntimeMount } from '../adapter.ts'
import { nativeHistory } from '../undo.ts'
import { jsonValueOf, localePack, workbookData } from '../wire.ts'

/**
 * Mount a workbook into one container.
 * @param input - the container and the wire snapshot to load.
 * @returns the live runtime behind the adapter.
 * @throws when the snapshot is not a workbook.
 */
export function mount(input: OfficeRuntimeMount): Promise<OfficeRuntime> {
  if (input.snapshot.kind !== 'workbook') {
    throw new Error('ui-office-editor: a document snapshot cannot mount in the sheets runtime')
  }
  const { container, snapshot } = input
  const data = workbookData(snapshot.data)
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: localePack(UniverSheetsCoreEnUS as Record<string, unknown>),
    },
    presets: [UniverSheetsCorePreset({ container })],
  })
  const workbook = univerAPI.createWorkbook(data)
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
    getSnapshot: () => ({ kind: 'workbook', data: jsonValueOf(workbook.save()) }),
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
