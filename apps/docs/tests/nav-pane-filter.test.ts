import { describe, expect, it } from 'vitest'
import {
  filterNavHeadings,
  navFilterNeedle,
  splitNavLabel,
} from '../src/renderer/components/nav-filter'

const headings = [
  { text: 'Chapter 1 Overview', level: 1 },
  { text: 'Introduction', level: 2 },
  { text: 'CHAPTER TWO', level: 1 },
  { text: '  Spaced  Heading  ', level: 3 },
]

describe('navigation pane search filter', () => {
  it('empty or blank query keeps every heading', () => {
    expect(filterNavHeadings(headings, '')).toBe(headings)
    expect(filterNavHeadings(headings, '   ')).toBe(headings)
    expect(filterNavHeadings(headings, '\t')).toBe(headings)
  })

  it('matches case-insensitively as a substring', () => {
    expect(filterNavHeadings(headings, 'chapter').map((h) => h.text)).toEqual([
      'Chapter 1 Overview',
      'CHAPTER TWO',
    ])
    expect(filterNavHeadings(headings, 'INTRO')).toHaveLength(1)
  })

  it('trims the query before matching', () => {
    expect(filterNavHeadings(headings, '  intro  ')).toHaveLength(1)
    expect(navFilterNeedle('  MiXeD  ')).toBe('mixed')
  })

  it('a non-matching query yields an empty list', () => {
    expect(filterNavHeadings(headings, 'zzz')).toEqual([])
  })

  it('splits a label around every case-insensitive occurrence', () => {
    expect(splitNavLabel('Chapter chart', 'CH')).toEqual([
      { text: 'Ch', hit: true },
      { text: 'apter ', hit: false },
      { text: 'ch', hit: true },
      { text: 'art', hit: false },
    ])
  })

  it('whole-string match and no-match shapes', () => {
    expect(splitNavLabel('intro', 'intro')).toEqual([{ text: 'intro', hit: true }])
    expect(splitNavLabel('intro', 'zzz')).toEqual([{ text: 'intro', hit: false }])
  })

  it('blank query returns one non-hit part', () => {
    expect(splitNavLabel('Introduction', '  ')).toEqual([{ text: 'Introduction', hit: false }])
  })
})
