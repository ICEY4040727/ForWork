/**
 * The Univer Docs runtime: the core preset, its English pack, and the CSS the
 * preset requires. Loaded only through the adapter's dynamic import.
 */
import { CommandType, LocaleType } from '@univerjs/core'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import UniverDocsCoreEnUS from '@univerjs/preset-docs-core/locales/en-US'
import '@univerjs/preset-docs-core/lib/index.css'
import { createUniver } from '@univerjs/presets'
import type { DocumentEditOperation } from '@deepseek-ai/dsh-session/src/document-events.ts'
import type { OfficeApplyOutcome } from '../command-bridge.ts'
import type { OfficeRuntime, OfficeRuntimeMount } from '../adapter.ts'
import { nativeHistory } from '../undo.ts'
import { documentData, jsonValueOf, localePack } from '../wire.ts'

/**
 * Refuse an operation the docs runtime cannot apply.
 * @param detail - the unsupported step or value.
 * @returns the refusal carrying the dropped content.
 */
function unsupported(detail: string): OfficeApplyOutcome {
  return { status: 'unsupported-feature', features: [{ kind: 'other', detail }] }
}

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
    apply(operation: DocumentEditOperation): OfficeApplyOutcome {
      const paragraphs = unit.getParagraphs()
      const targets: { paragraph: (typeof paragraphs)[number]; text: string }[] = []
      for (const step of operation) {
        if (step.kind !== 'paragraph') return unsupported(step.kind)
        const paragraph = paragraphs[step.index]
        if (paragraph === undefined) return unsupported(`paragraph ${step.index}`)
        targets.push({ paragraph, text: step.text })
      }
      for (const target of targets) target.paragraph.setText(target.text)
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
