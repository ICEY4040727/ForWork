/**
 * The Univer Sheets runtime: the core preset, its English pack, and the CSS
 * the preset requires. Loaded only through the editor runtime chunk.
 */
import { CommandType, LocaleType } from '@univerjs/core'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver } from '@univerjs/presets'
import type { DocumentEditOperation } from '@deepseek-ai/dsh-session/src/document-events.ts'
import type { OfficeApplyOutcome } from '../command-bridge.ts'
import type { OfficeRuntime, OfficeRuntimeMount } from '../adapter.ts'
import { nativeHistory } from '../undo.ts'
import { jsonValueOf, localePack, workbookData } from '../wire.ts'

/**
 * Refuse an operation the sheets runtime cannot apply.
 * @param detail - the unsupported step or value.
 * @returns the refusal carrying the dropped content.
 */
function unsupported(detail: string): OfficeApplyOutcome {
  return { status: 'unsupported-feature', features: [{ kind: 'other', detail }] }
}

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
    apply(operation: DocumentEditOperation): OfficeApplyOutcome {
      const cells: {
        sheet: NonNullable<ReturnType<typeof workbook.getSheetBySheetId>>
        row: number
        column: number
        value: string | number | boolean | null
      }[] = []
      for (const step of operation) {
        if (step.kind !== 'cell') return unsupported(step.kind)
        const { value } = step
        if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
          return unsupported(`cell value ${Array.isArray(value) ? 'array' : typeof value}`)
        }
        const sheet = workbook.getSheetBySheetId(step.sheetId)
        if (sheet === null) return unsupported(`sheet ${step.sheetId}`)
        cells.push({ sheet, row: step.row, column: step.column, value })
      }
      for (const cell of cells) {
        const range = cell.sheet.getRange(cell.row, cell.column)
        if (cell.value === null) range.setValue({ v: null, f: null, p: null })
        else range.setValue(cell.value)
      }
      return { status: 'applied' }
    },
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
