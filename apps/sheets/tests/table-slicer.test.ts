import { describe, expect, it } from 'vitest'

import {
  TABLE_SLICER_MAX_MEMBERS,
  tableSlicerMembers,
  tableSlicerSelection,
} from '../src/renderer/pivot-actions'

describe('tableSlicerMembers', () => {
  it('collects distinct values in first-appearance order and collapses blanks', () => {
    const members = tableSlicerMembers(
      ['de', 'fr', 'de', null, '', 10, true, 'fr', false],
      '(blank)',
    )
    expect(members.map((member) => member.label)).toEqual([
      'de',
      'fr',
      '(blank)',
      '10',
      'TRUE',
      'FALSE',
    ])
    expect(members.map((member) => member.member)).toEqual([0, 1, 2, 3, 4, 5])
  })

  it('caps the member list', () => {
    const values = Array.from({ length: TABLE_SLICER_MAX_MEMBERS + 50 }, (_v, i) => `v${i}`)
    expect(tableSlicerMembers(values, '(blank)')).toHaveLength(TABLE_SLICER_MAX_MEMBERS)
  })
})

describe('tableSlicerSelection', () => {
  const members = tableSlicerMembers(['de', 'fr', 'it'], '(blank)')

  it('selects everything without criteria', () => {
    expect(tableSlicerSelection(members, null)).toEqual([0, 1, 2])
    expect(tableSlicerSelection(members, undefined)).toEqual([0, 1, 2])
    expect(tableSlicerSelection(members, {})).toEqual([0, 1, 2])
  })

  it('selects exactly the values the filter keeps', () => {
    // OOXML <filters> lists the KEPT values.
    expect(tableSlicerSelection(members, { values: ['de', 'it'] })).toEqual([0, 2])
    expect(tableSlicerSelection(members, { values: [], blank: true })).toEqual([])
  })

  it('maps a kept blank criterion onto the blank member', () => {
    const withBlank = tableSlicerMembers(['de', '', 'it'], '(blank)')
    // The blank member's label is the locale's blank text, matched by the
    // blank flag, not by the label.
    expect(tableSlicerSelection(withBlank, { values: ['de'], blank: true })).toEqual([0, 1])
    expect(tableSlicerSelection(withBlank, { values: ['de'] })).toEqual([0])
  })

  it('starts fully selected on custom-criteria filters (unrepresentable)', () => {
    // A criteria object without values/blank carries no member mapping.
    expect(tableSlicerSelection(members, {})).toEqual([0, 1, 2])
  })
})
