import type JSZip from 'jszip'
import type { Background, Color, Deck, LayoutId, Paragraph, PlaceholderRole, Slide, SlideElement, Stroke, TableCell, TableElement, TextBody, Theme } from '../deck.ts'
import { EMU_PER_POINT, LAYOUTS, SLIDE_SIZES } from '../deck.ts'
import { coverCrop, describeElement } from '../elements.ts'
import { LAYOUT_NAMES } from '../layouts.ts'
import { cellUnder } from '../tables.ts'
import { isSlot } from '../themes.ts'
import { bulletFor, numberingFor, paragraphIndent } from '../text.ts'
import { placeholderIndex, placeholderNames, placeholderSlots } from './placeholders.ts'
import { child, childrenNamed, descendants, elements, find, parseXml, serializeXml, xml, type XmlElement, type XmlNode } from './xml.ts'

/*
 * What a PowerPoint file needs after PptxGenJS has written it. PptxGenJS repeats a paragraph's
 * settings before each of its runs, takes a placeholder's position and text box settings from its
 * layout whatever the slide says, cannot write gradients, crops, picture outlines, adjust values
 * or a table's cells as PowerPoint has them, numbers tables apart from other shapes, and leaves
 * Office's colours in the theme; so each slide's elements get their paragraph and text box
 * settings, positions, geometry and table cells from the deck itself, every shape a unique id,
 * empty placeholders Herald did not have go, and the theme gets the deck's colours.
 */

const emu = (points: number): string => String(Math.round(points * EMU_PER_POINT))

/** A colour element: a theme slot by name or a literal, with its opacity. */
export function colorXml(color: Color, alpha?: number): XmlElement {
  const opacity = alpha !== undefined && alpha < 1 ? [xml('a:alpha', { val: Math.round(Math.max(0, alpha) * 100000) })] : []

  return isSlot(color) ? xml('a:schemeClr', { val: color }, opacity) : xml('a:srgbClr', { val: color.slice(1).toUpperCase() }, opacity)
}

const solid = (color: Color, alpha?: number): XmlElement => xml('a:solidFill', {}, [colorXml(color, alpha)])

const DASH = { solid: 'solid', dash: 'dash', dot: 'sysDot', dashDot: 'dashDot', longDash: 'lgDash' } as const

function lineXml(stroke: Stroke | null): XmlElement {
  if (!stroke || stroke.width <= 0) {
    return xml('a:ln', {}, [xml('a:noFill')])
  }

  return xml('a:ln', { w: emu(stroke.width) }, [solid(stroke.color, stroke.alpha), xml('a:prstDash', { val: DASH[stroke.dash] })])
}

const ALIGN = { left: 'l', center: 'ctr', right: 'r', justify: 'just' } as const

/** A paragraph's settings as PowerPoint reads them, all spelled out so a layout's defaults never apply. */
export function paragraphXml(paragraph: Paragraph): XmlElement {
  const { margin, indent } = paragraphIndent(paragraph)
  const level = paragraph.level ?? 0
  const points = (name: string, value: number | undefined) => xml(name, {}, [xml('a:spcPts', { val: Math.round((value ?? 0) * 100) })])
  const bullet: XmlNode[] =
    paragraph.list === 'bullet'
      ? [xml('a:buClrTx'), xml('a:buSzTx'), xml('a:buFontTx'), xml('a:buChar', { char: paragraph.bullet || bulletFor(level) })]
      : paragraph.list === 'number'
        ? [xml('a:buClrTx'), xml('a:buSzTx'), xml('a:buFontTx'), xml('a:buAutoNum', { type: paragraph.numbering ?? numberingFor(level), startAt: paragraph.startAt && paragraph.startAt !== 1 ? paragraph.startAt : undefined })]
        : [xml('a:buNone')]

  return xml('a:pPr', { marL: emu(margin), indent: emu(indent), algn: ALIGN[paragraph.align ?? 'left'], lvl: level || undefined }, [
    xml('a:lnSpc', {}, [xml('a:spcPct', { val: Math.round((paragraph.lineSpacing ?? 1) * 100000) })]),
    points('a:spcBef', paragraph.spaceBefore),
    points('a:spcAft', paragraph.spaceAfter),
    ...bullet
  ])
}

/** A text box's settings; `shrink` is the factor its text shrinks to, as last drawn (PowerPoint shows it so until the text is edited). */
export function bodyPropertiesXml(body: TextBody, shrink?: number): XmlElement {
  const [left, top, right, bottom] = body.inset
  const scale = shrink !== undefined && shrink < 1 ? Math.round(shrink * 100000) : undefined
  const fit = body.fit === 'shrink' ? [xml('a:normAutofit', { fontScale: scale })] : body.fit === 'grow' ? [xml('a:spAutoFit')] : []

  return xml('a:bodyPr', { wrap: body.wrap ? 'square' : 'none', lIns: emu(left), tIns: emu(top), rIns: emu(right), bIns: emu(bottom), rtlCol: '0', anchor: body.anchor === 'middle' ? 'ctr' : body.anchor === 'bottom' ? 'b' : 't' }, fit)
}

/** A text body's paragraphs with their own settings, and the body's settings, in place of PptxGenJS's. */
function finishText(txBody: XmlElement, body: TextBody, shrink?: number): void {
  const paragraphs = childrenNamed(txBody, 'a:p')

  txBody.children = txBody.children.map((node) => (typeof node !== 'string' && node.name === 'a:bodyPr' ? bodyPropertiesXml(body, shrink) : node))

  paragraphs.forEach((p, index) => {
    const model = body.paragraphs[Math.min(index, body.paragraphs.length - 1)]
    p.children = [paragraphXml(model), ...p.children.filter((node) => typeof node === 'string' || node.name !== 'a:pPr')]
  })
}

function setTransform(spPr: XmlElement, element: SlideElement): void {
  const attrs: Record<string, string | undefined> = { rot: element.rotation ? String(Math.round(element.rotation * 60000)) : undefined, flipH: element.flipH ? '1' : undefined, flipV: element.flipV ? '1' : undefined }
  const transform = xml('a:xfrm', attrs, [xml('a:off', { x: emu(element.x), y: emu(element.y) }), xml('a:ext', { cx: emu(element.width), cy: emu(element.height) })])
  const at = spPr.children.findIndex((node) => typeof node !== 'string' && node.name === 'a:xfrm')

  if (at >= 0) {
    spPr.children[at] = transform
  } else {
    spPr.children.unshift(transform)
  }
}

/** Put `line` in a shape's properties where the schema wants it: after the geometry and the fill. */
function setLine(spPr: XmlElement, line: XmlElement): void {
  const at = spPr.children.findIndex((node) => typeof node !== 'string' && node.name === 'a:ln')

  if (at >= 0) {
    spPr.children[at] = line

    return
  }

  const after = ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']
  let index = -1
  spPr.children.forEach((node, position) => {
    if (typeof node !== 'string' && after.includes(node.name)) {
      index = position
    }
  })
  spPr.children.splice(index + 1, 0, line)
}

/** How much a text body's text was shrunk to fit when last drawn, when the window knows. */
export type ShrinkOf = (body: TextBody) => number | undefined

/** The name PowerPoint's selection pane shows: the element's own, or what it is and a number (Herald's ids stay in Herald's own copy). */
function nameOf(element: SlideElement, names: Map<string, number>): string {
  const base = element.name ?? describeElement(element)
  const n = (names.get(base) ?? 0) + 1
  names.set(base, n)

  return element.name ?? `${base} ${n}`
}

function finishElement(node: XmlElement, element: SlideElement, names: Map<string, number>, shrink?: ShrinkOf): void {
  const nv = child(node, node.name === 'p:pic' ? 'p:nvPicPr' : 'p:nvSpPr')
  const cNvPr = child(nv, 'p:cNvPr')

  if (cNvPr) {
    cNvPr.attrs.name = nameOf(element, names)
  }

  const spPr = child(node, 'p:spPr')

  if (spPr) {
    setTransform(spPr, element)
  }

  if (element.kind === 'shape' && element.adjust && spPr) {
    const list = find(spPr, 'a:prstGeom/a:avLst')

    if (list) {
      list.children = Object.entries(element.adjust).map(([name, value]) => xml('a:gd', { name, fmla: `val ${Math.round(value)}` }))
    }
  }

  if ((element.kind === 'text' || element.kind === 'shape') && spPr) {
    setLine(spPr, lineXml(element.stroke))
  }

  if (element.kind === 'image') {
    if (spPr && element.stroke) {
      setLine(spPr, lineXml(element.stroke))
    }

    const fill = child(node, 'p:blipFill')

    if (fill && element.crop) {
      const crop = element.crop
      const rect = xml('a:srcRect', { l: Math.round(crop.left * 100000), t: Math.round(crop.top * 100000), r: Math.round(crop.right * 100000), b: Math.round(crop.bottom * 100000) })
      fill.children = fill.children.filter((entry) => typeof entry === 'string' || entry.name !== 'a:srcRect')
      const blip = fill.children.findIndex((entry) => typeof entry !== 'string' && entry.name === 'a:blip')
      fill.children.splice(blip + 1, 0, rect)
    }
  }

  const txBody = child(node, 'p:txBody')

  if (txBody && (element.kind === 'text' || element.kind === 'shape')) {
    finishText(txBody, element.body, shrink?.(element.body))
  }
}

/** Lengths in EMU, each rounded where it ends, so together they are their total rounded. */
function emuSpans(lengths: readonly number[]): number[] {
  let at = 0
  let written = 0

  return lengths.map((length) => {
    at += length
    const end = Math.round(at * EMU_PER_POINT)
    const span = end - written
    written = end

    return span
  })
}

const ANCHOR = { top: 't', middle: 'ctr', bottom: 'b' } as const

/** A cell's settings as PowerPoint reads them: its margins and anchor, the table's lines on every side, and its fill. */
function cellPropertiesXml(cell: TableCell, stroke: Stroke | null): XmlElement {
  const [left, top, right, bottom] = cell.body.inset
  const side = (name: string): XmlElement =>
    stroke && stroke.width > 0 ? xml(name, { w: emu(stroke.width), cap: 'flat', cmpd: 'sng', algn: 'ctr' }, [solid(stroke.color, stroke.alpha), xml('a:prstDash', { val: DASH[stroke.dash] })]) : xml(name, {}, [xml('a:noFill')])

  return xml('a:tcPr', { marL: emu(left), marR: emu(right), marT: emu(top), marB: emu(bottom), anchor: ANCHOR[cell.body.anchor] }, [
    side('a:lnL'),
    side('a:lnR'),
    side('a:lnT'),
    side('a:lnB'),
    cell.fill ? solid(cell.fill.color, cell.fill.alpha) : xml('a:noFill')
  ])
}

/** A cell's text as a table holds it: empty body properties (the cell's own give its margins), a list style, and at least one paragraph with its settings. */
function cellTextXml(txBody: XmlElement | undefined, body: TextBody): XmlElement {
  const written = childrenNamed(txBody, 'a:p')
  const paragraphs = written.length ? written : [xml('a:p', {}, [xml('a:endParaRPr', { lang: 'en-US', dirty: '0' })])]

  paragraphs.forEach((p, index) => {
    const model = body.paragraphs[Math.min(index, body.paragraphs.length - 1)]
    p.children = [paragraphXml(model), ...p.children.filter((node) => typeof node === 'string' || node.name !== 'a:pPr')]
  })

  return xml('a:txBody', {}, [xml('a:bodyPr'), xml('a:lstStyle'), ...paragraphs])
}

/**
 * A table as the deck has it: its place, a grid column for each column, and each row with a cell
 * for every column, each cell its text and then its settings, a merged cell's reach as gridSpan and
 * rowSpan where it starts and the cells it covers marked hMerge and vMerge, as PowerPoint writes them.
 * PptxGenJS wrote every cell as one of its own, so its cells are the deck's one for one.
 */
function finishTable(frame: XmlElement, table: TableElement): void {
  const widths = emuSpans(table.columns)
  const heights = emuSpans(table.rows)
  const transform = child(frame, 'p:xfrm')
  const tbl = find(frame, 'a:graphic/a:graphicData/a:tbl')

  if (transform) {
    transform.attrs = {}
    transform.children = [xml('a:off', { x: emu(table.x), y: emu(table.y) }), xml('a:ext', { cx: widths.reduce((sum, width) => sum + width, 0), cy: heights.reduce((sum, height) => sum + height, 0) })]
  }

  if (!tbl) {
    return
  }

  const rows = childrenNamed(tbl, 'a:tr')
  tbl.children = [
    child(tbl, 'a:tblPr') ?? xml('a:tblPr'),
    xml(
      'a:tblGrid',
      {},
      widths.map((width) => xml('a:gridCol', { w: width }))
    ),
    ...table.cells.map((row, r) => {
      const written = childrenNamed(rows[r], 'a:tc')

      return xml(
        'a:tr',
        { h: heights[r] },
        row.map((cell, c) => {
          const under = cell.merged ? cellUnder(table, { row: r, column: c }) : null
          const reach = under ? { hMerge: c > under.column ? '1' : undefined, vMerge: r > under.row ? '1' : undefined } : { gridSpan: cell.colSpan, rowSpan: cell.rowSpan }

          return xml('a:tc', reach, [cellTextXml(child(written[c], 'a:txBody'), cell.body), cellPropertiesXml(cell, table.stroke)])
        })
      )
    })
  ]
}

/** Every shape on a slide with an id of its own, as PowerPoint requires: PptxGenJS numbers tables apart from the rest. */
function uniqueIds(tree: XmlElement): void {
  const named = descendants(tree, 'p:cNvPr')
  const seen = new Set<string>()
  let next = Math.max(0, ...named.map((node) => Number(node.attrs.id) || 0)) + 1

  for (const node of named) {
    if (!/^\d+$/.test(node.attrs.id ?? '') || seen.has(node.attrs.id)) {
      node.attrs.id = String(next++)
    }

    seen.add(node.attrs.id)
  }
}

function backgroundXml(background: Background & { kind: 'gradient' }): XmlElement {
  const stops = [...background.stops].sort((a, b) => a.at - b.at).map((stop) => xml('a:gs', { pos: Math.round(stop.at * 100000) }, [colorXml(stop.color)]))
  const angle = (((background.angle % 360) + 360) % 360) * 60000

  return xml('p:bg', {}, [xml('p:bgPr', {}, [xml('a:gradFill', { rotWithShape: '1' }, [xml('a:gsLst', {}, stops), xml('a:lin', { ang: Math.round(angle), scaled: '0' })]), xml('a:effectLst')])])
}

function finishBackground(cSld: XmlElement, slide: Slide, deck: Deck): void {
  const background = slide.background

  if (background?.kind === 'gradient') {
    const at = cSld.children.findIndex((node) => typeof node !== 'string' && node.name === 'p:bg')
    const replacement = backgroundXml(background)

    if (at >= 0) {
      cSld.children[at] = replacement
    } else {
      cSld.children.unshift(replacement)
    }
  } else if (background?.kind === 'image') {
    const rect = find(cSld, 'p:bg/p:bgPr/a:blipFill/a:srcRect')
    const crop = coverCrop(background.natural, deck.size)

    if (rect && crop) {
      rect.attrs = { l: String(Math.round(crop.left * 100000)), t: String(Math.round(crop.top * 100000)), r: String(Math.round(crop.right * 100000)), b: String(Math.round(crop.bottom * 100000)) }
    }
  }
}

function finishSlide(root: XmlElement, slide: Slide, deck: Deck, shrink?: ShrinkOf): void {
  const cSld = child(root, 'p:cSld')
  const tree = child(cSld, 'p:spTree')

  if (!cSld || !tree) {
    return
  }

  const names = new Map<string, number>()
  const byId = new Map(slide.elements.map((element) => [element.id, element]))
  const byIndex = new Map([...placeholderSlots(slide)].map(([id, name]) => [String(placeholderIndex(slide.layout, name)), id]))

  tree.children = tree.children.filter((node) => {
    if (typeof node !== 'string' && node.name === 'p:graphicFrame') {
      const cNvPr = find(node, 'p:nvGraphicFramePr/p:cNvPr')
      const element = byId.get(cNvPr?.attrs.name ?? '')

      if (cNvPr && element?.kind === 'table') {
        cNvPr.attrs.name = nameOf(element, names)
        finishTable(node, element)
      }

      return true
    }

    if (typeof node === 'string' || (node.name !== 'p:sp' && node.name !== 'p:pic')) {
      return true
    }

    const nv = child(node, node.name === 'p:pic' ? 'p:nvPicPr' : 'p:nvSpPr')
    const ph = find(nv, 'p:nvPr/p:ph')
    const id = ph ? byIndex.get(ph.attrs.idx ?? '') : child(nv, 'p:cNvPr')?.attrs.name

    // A layout placeholder PptxGenJS added that the slide does not have.
    if (ph && !id) {
      return false
    }

    const element = id ? byId.get(id) : undefined

    if (element) {
      finishElement(node, element, names, shrink)

      if (ph && element.placeholder) {
        ph.attrs.type = placeholderType(slide.layout, element.placeholder.role, ph.attrs.type)
      }
    }

    return true
  })

  uniqueIds(tree)
  finishBackground(cSld, slide, deck)

  root.children = root.children.filter((node) => typeof node === 'string' || node.name !== 'p:transition')

  if (deck.transition !== 'none') {
    const transition = xml('p:transition', { spd: 'med' }, [deck.transition === 'fade' ? xml('p:fade') : xml('p:push', { dir: 'l' })])
    const timing = root.children.findIndex((node) => typeof node !== 'string' && (node.name === 'p:timing' || node.name === 'p:extLst'))
    root.children.splice(timing < 0 ? root.children.length : timing, 0, transition)
  }
}

/** Speaker notes a paragraph a line (PptxGenJS puts them in one). */
function finishNotes(root: XmlElement, notes: string): void {
  for (const sp of elements(find(root, 'p:cSld/p:spTree')).filter((node) => node.name === 'p:sp')) {
    if (find(sp, 'p:nvSpPr/p:nvPr/p:ph')?.attrs.type !== 'body') {
      continue
    }

    const txBody = child(sp, 'p:txBody')

    if (txBody) {
      const lines = notes.split('\n')
      txBody.children = [
        ...txBody.children.filter((node) => typeof node !== 'string' && (node.name === 'a:bodyPr' || node.name === 'a:lstStyle')),
        ...lines.map((line) => (line ? xml('a:p', {}, [xml('a:r', {}, [xml('a:rPr', { lang: 'en-US', dirty: '0' }), xml('a:t', {}, [line])])]) : xml('a:p', {}, [xml('a:endParaRPr', { lang: 'en-US', dirty: '0' })])))
      ]
    }
  }
}

const SCHEME_ORDER: [string, keyof Theme['colors']][] = [
  ['a:dk1', 'tx1'],
  ['a:lt1', 'bg1'],
  ['a:dk2', 'tx2'],
  ['a:lt2', 'bg2'],
  ['a:accent1', 'accent1'],
  ['a:accent2', 'accent2'],
  ['a:accent3', 'accent3'],
  ['a:accent4', 'accent4'],
  ['a:accent5', 'accent5'],
  ['a:accent6', 'accent6'],
  ['a:hlink', 'accent1'],
  ['a:folHlink', 'accent5']
]

/** The theme's colours and name in place of Office's (its fonts PptxGenJS sets itself). */
function finishTheme(root: XmlElement, theme: Theme): void {
  root.attrs.name = theme.name
  const elementsOf = child(root, 'a:themeElements')
  const scheme = xml(
    'a:clrScheme',
    { name: theme.name },
    SCHEME_ORDER.map(([name, slot]) => xml(name, {}, [xml('a:srgbClr', { val: theme.colors[slot].slice(1).toUpperCase() })]))
  )

  if (elementsOf) {
    elementsOf.children = elementsOf.children.map((node) => (typeof node !== 'string' && node.name === 'a:clrScheme' ? scheme : node))
  }
}

async function rewrite(zip: JSZip, name: string, change: (root: XmlElement) => void): Promise<void> {
  const file = zip.file(name)

  if (!file) {
    return
  }

  const root = parseXml(await file.async('string'), { canonical: false })
  change(root)
  zip.file(name, serializeXml(root))
}

/** PresentationML's name for each layout, which tells other apps (and readers) what the layout is for. */
const LAYOUT_TYPES: Record<LayoutId, string> = {
  title: 'title',
  'title-content': 'obj',
  'two-content': 'twoObj',
  section: 'secHead',
  'title-only': 'titleOnly',
  blank: 'blank',
  'picture-caption': 'picTx',
  comparison: 'twoTxTwoObj'
}

/** A title slide's placeholders are a centred title and a subtitle, as PowerPoint's own are; PptxGenJS knows only titles and bodies. */
function placeholderType(layout: LayoutId, role: PlaceholderRole, written: string | undefined): string {
  if (layout === 'title' && role === 'title') {
    return 'ctrTitle'
  }

  if (layout === 'title' && role === 'subtitle') {
    return 'subTitle'
  }

  return written ?? 'body'
}

/** Finish a PowerPoint file PptxGenJS wrote for `deck`. */
export async function finishPresentation(zip: JSZip, deck: Deck, shrink?: ShrinkOf): Promise<void> {
  const byName = new Map(LAYOUTS.map((layout) => [LAYOUT_NAMES[layout], layout]))

  for (const name of Object.keys(zip.files).filter((entry) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(entry))) {
    await rewrite(zip, name, (root) => {
      const layout = byName.get(child(root, 'p:cSld')?.attrs.name ?? '')

      if (!layout) {
        return
      }

      root.attrs.type = LAYOUT_TYPES[layout]
      const roles = new Map(placeholderNames(layout).map((entry) => [String(placeholderIndex(layout, entry.name)), entry.role]))

      for (const ph of descendants(find(root, 'p:cSld/p:spTree'), 'p:ph')) {
        const role = roles.get(ph.attrs.idx ?? '')

        if (role) {
          ph.attrs.type = placeholderType(layout, role, ph.attrs.type)
        }
      }
    })
  }

  for (const [index, slide] of deck.slides.entries()) {
    await rewrite(zip, `ppt/slides/slide${index + 1}.xml`, (root) => finishSlide(root, slide, deck, shrink))

    if (slide.notes) {
      await rewrite(zip, `ppt/notesSlides/notesSlide${index + 1}.xml`, (root) => finishNotes(root, slide.notes))
    }
  }

  await rewrite(zip, 'ppt/theme/theme1.xml', (root) => finishTheme(root, deck.theme))

  const app = zip.file('docProps/app.xml')

  if (app) {
    const format = deck.size.width === SLIDE_SIZES.standard.width && deck.size.height === SLIDE_SIZES.standard.height ? 'On-screen Show (4:3)' : deck.size.width / deck.size.height === 16 / 9 ? 'On-screen Show (16:9)' : 'Custom'
    zip.file('docProps/app.xml', (await app.async('string')).replace(/<PresentationFormat>[^<]*<\/PresentationFormat>/, `<PresentationFormat>${format}</PresentationFormat>`))
  }

  // PptxGenJS gives a content type to a slide master for every slide, though it writes only one.
  await rewrite(zip, '[Content_Types].xml', (root) => {
    root.children = root.children.filter((node) => typeof node === 'string' || node.name !== 'Override' || zip.file((node.attrs.PartName ?? '').slice(1)) !== null)
  })
}
