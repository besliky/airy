import type { zh } from './zh'

export const en = {
  reviewCompareMerge: 'Compare (Legal Blackline)',
  reviewComparePanel: 'Show Differences Pane Only',
  reviewCompareMerged:
    'Compared with {name}: merged {added} insertions, {removed} deletions and {changed} changes as tracked changes',
  reviewCompareIdentical: 'No differences with {name}: the documents are identical',
} satisfies Record<keyof typeof zh, string>
