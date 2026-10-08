import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { findSlide } from '../deck.ts'
import * as model from '../model.ts'
import { settleSpans, withCell } from '../tables.ts'
import { DEFAULT_THEME } from '../themes.ts'
import { SlideView } from './SlideView.tsx'

describe('a table on a slide', () => {
  it('draws a grid of cells with their text, fills and lines, a merged cell spanning the cells it covers', () => {
    const start = model.newDeck('Plan')
    const slideId = start.slides[0].id
    const added = model.addTable(start, slideId, { rows: 2, columns: 3, cells: [['Name', 'Role', 'Team']] })
    const { deck } = model.updateElements(
      added.deck,
      slideId,
      [added.elementId],
      (element) => {
        if (element.kind !== 'table') {
          return element
        }

        const spanning = withCell(element, { row: 1, column: 1 }, (cell) => ({ ...cell, colSpan: 2 }))

        return { ...spanning, cells: settleSpans(spanning.cells, spanning.columns.length) }
      },
      'Merge'
    )
    const html = renderToStaticMarkup(createElement(SlideView, { deck, slide: findSlide(deck, slideId)!, scale: 1, mode: 'present' }))

    expect(html).toContain(`data-element-id="${added.elementId}"`)
    expect(html).toContain('grid-template-columns:240px 240px 240px')
    expect(html).toContain('grid-template-rows:minmax(29.2px, auto) minmax(29.2px, auto)')
    expect(html.match(/class="hs-cell"/g)).toHaveLength(5)
    expect(html).toContain('data-row="1" data-column="1" data-anchor="top" style="grid-row:2 / span 1;grid-column:2 / span 2')
    expect(html).not.toContain('data-row="1" data-column="2"')
    expect(html).toContain(`background-color:${DEFAULT_THEME.colors.accent1};outline:1px solid ${DEFAULT_THEME.colors.bg1};outline-offset:-0.5px`)
    expect(html).toMatch(/<span[^>]*>Name<\/span>/)
  })

  it('draws the text editor in the cell being typed into, and the other cells as they are', () => {
    const start = model.newDeck('Plan')
    const slideId = start.slides[0].id
    const { deck, elementId } = model.addTable(start, slideId, { rows: 2, columns: 2, cells: [['a', 'b'], ['c', 'd']] })
    const editing = { id: elementId, cell: { row: 1, column: 0 }, render: () => createElement('i', { 'data-typing': '' }) }
    const html = renderToStaticMarkup(createElement(SlideView, { deck, slide: findSlide(deck, slideId)!, scale: 1, mode: 'edit', editing }))

    expect(html).toMatch(/data-row="1" data-column="0"[^>]*><i data-typing="">/)
    expect(html.match(/data-typing/g)).toHaveLength(1)
    expect(html).not.toMatch(/<span[^>]*>c<\/span>/)
    expect(html).toMatch(/<span[^>]*>d<\/span>/)
  })
})
