/**
 * Length caps for user-provided file/tab names must count CODE POINTS, not
 * UTF-16 code units: `String.prototype.slice` cuts surrogate pairs in half,
 * leaving a lone surrogate — an unpaired code unit that cannot be encoded in
 * UTF-8 (WTF-8 aside) and that NTFS and macOS filesystems reject or mangle,
 * so the "safe" capped name failed exactly where it had to work (BUG-412: a
 * name whose 80th code unit was the low half of an emoji produced a file the
 * OS refused to create). Splitting here never splits a pair: the cap counts
 * whole code points.
 */
export function truncateByCodePoints(value: string, maxCodePoints: number): string {
  if (!Number.isInteger(maxCodePoints) || maxCodePoints < 1)
    throw new RangeError('maxCodePoints must be a positive integer')
  // UTF-16 length is an upper bound of the code-point count, so a string this
  // short needs no work (and spreading huge strings for nothing would show up
  // on hot paths)
  if (value.length <= maxCodePoints) return value
  let out = ''
  let kept = 0
  // for..of iterates code points (surrogate pairs stay whole)
  for (const char of value) {
    if (kept === maxCodePoints) break
    out += char
    kept += 1
  }
  return out
}
