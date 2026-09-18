import type { zh } from './zh'

export const en = {
  reviewCompareMerge: 'Compare (Legal Blackline)',
  reviewCompareMergeDesc: 'Show the differences as tracked changes you can accept or reject',
  reviewComparePanel: 'Show Differences Pane Only',
  reviewComparePanelDesc:
    'List the paragraph differences in a side pane without editing the document',
  reviewCompareMerged:
    'Compared with {name}: merged {added} insertions, {removed} deletions and {changed} changes as tracked changes',
  reviewCompareMergedApprox:
    'Compared with {name}: merged {added} insertions, {removed} deletions and {changed} changes as tracked changes (documents too large for exact paragraph matching)',
  reviewCompareIdentical: 'No differences with {name}: the documents are identical',
  reviewCompareDegraded:
    'The documents are too large for exact paragraph matching: differences were paired by position',
  reviewComparePendingRevisions:
    'The document has pending tracked changes. Accept or reject them before comparing again',
  reviewCompareReadonly:
    'Compare (Legal Blackline) needs an editable document; this document is read-only',
  reviewComparing: 'Comparing…',
} satisfies Record<keyof typeof zh, string>
