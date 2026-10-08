import { Extension } from '@tiptap/core'
import { type EditorState, Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'

/*
 * The text a request to Hermes is about: what was selected when the person asked, kept through
 * every edit made meanwhile (theirs or Hermes's) and tinted in the page while Hermes works on it.
 * Hermes's commands write to it with `at: "marked"`, so the answer lands where it was asked for even
 * after the person has clicked elsewhere.
 */

export interface MarkedRange {
  from: number
  to: number
}

interface MarkedState {
  range: MarkedRange | null
  decorations: DecorationSet
}

export const markedKey = new PluginKey<MarkedState>('docsHermesMarked')

const EMPTY: MarkedState = { range: null, decorations: DecorationSet.empty }

function stateFor(state: { doc: EditorState['doc'] }, range: MarkedRange | null): MarkedState {
  if (!range || range.to < range.from) {
    return EMPTY
  }

  const decorations = range.to > range.from ? DecorationSet.create(state.doc, [Decoration.inline(range.from, range.to, { class: 'docs-hermes-marked' })]) : DecorationSet.empty

  return { range, decorations }
}

export function markedPlugin(): Plugin<MarkedState> {
  return new Plugin<MarkedState>({
    key: markedKey,
    state: {
      init: () => EMPTY,
      apply(tr, value) {
        const meta = tr.getMeta(markedKey) as { range: MarkedRange | null } | undefined

        if (meta) {
          return stateFor(tr, meta.range)
        }

        if (!value.range || !tr.docChanged) {
          return value
        }

        // Text typed at either edge joins neither side: the range stays what was asked about.
        const range = { from: tr.mapping.map(value.range.from, 1), to: tr.mapping.map(value.range.to, -1) }

        return stateFor(tr, range.to >= range.from ? range : { from: range.from, to: range.from })
      }
    },
    props: {
      decorations: (state) => markedKey.getState(state)?.decorations
    }
  })
}

export const HermesMarked = Extension.create({
  name: 'hermesMarked',
  addProseMirrorPlugins() {
    return [markedPlugin()]
  }
})

export const markedRangeOf = (state: EditorState): MarkedRange | null => markedKey.getState(state)?.range ?? null

/** Mark a range (the selection, by default) for Hermes, or clear the mark with null. */
export function setMarked(view: EditorView, range: MarkedRange | null | 'selection' = 'selection'): MarkedRange | null {
  const chosen = range === 'selection' ? { from: view.state.selection.from, to: view.state.selection.to } : range
  view.dispatch(view.state.tr.setMeta(markedKey, { range: chosen }).setMeta('addToHistory', false))

  return chosen
}
