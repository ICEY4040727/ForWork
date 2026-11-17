/**
 * The Univer Sheets runtime: the core preset, its English pack, and the CSS
 * the preset requires. Loaded only through the editor runtime chunk.
 */
import { CommandType, LocaleType } from '@univerjs/core'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver } from '@univerjs/presets'
import type { DocumentEditOperation } from '@forwork/session/src/document-events.ts'
import type { OfficeApplyOutcome } from '../command-bridge.ts'
import type { OfficeRuntime, OfficeRuntimeMount } from '../adapter.ts'
import { nativeHistory } from '../undo.ts'
import { jsonValueOf, localePack, workbookData } from '../wire.ts'
import { applySheetOperation } from './sheet-operations.ts'
import type { SheetWorkbook } from './sheet-operations.ts'

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
  const sheets: SheetWorkbook = {
    getSheetBySheetId(sheetId) {
      const sheet = workbook.getSheetBySheetId(sheetId)
      if (sheet === null) return null
      return {
        getMaxRows: () => sheet.getMaxRows(),
        getMaxColumns: () => sheet.getMaxColumns(),
        insertRows: (row, count) => { sheet.insertRows(row, count) },
        deleteRows: (row, count) => { sheet.deleteRows(row, count) },
        insertColumns: (column, count) => { sheet.insertColumns(column, count) },
        deleteColumns: (column, count) => { sheet.deleteColumns(column, count) },
        getRange: (startRow, startColumn, endRow, endColumn) =>
          sheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1),
      }
    },
  }
  return Promise.resolve({
    getSnapshot: () => ({ kind: 'workbook', data: jsonValueOf(workbook.save()) }),
    apply: (operation: DocumentEditOperation): OfficeApplyOutcome => applySheetOperation(sheets, operation),
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
