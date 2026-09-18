import type { zh } from './zh'

export const en = {
  reviewCompareMerge: 'Compare (Legal Blackline)',
  reviewComparePanel: 'Show Differences Pane Only',
  reviewCompareMerged:
    'Compared with {name}: merged {added} insertions, {removed} deletions and {changed} changes as tracked changes',
  reviewCompareMergedApprox:
    'Compared with {name}: merged {added} insertions, {removed} deletions and {changed} changes as tracked changes (documents too large for exact paragraph matching)',
  reviewCompareIdentical: 'No differences with {name}: the documents are identical',
  reviewCompareDegraded:
    'The documents are too large for exact paragraph matching: differences were paired by position',
} satisfies Record<keyof typeof zh, string>
