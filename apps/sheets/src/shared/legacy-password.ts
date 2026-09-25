/**
 * Excel's legacy worksheet/workbook password hash (the `password=` attribute
 * of <sheetProtection> / <workbookProtection>). The scheme is documented
 * public knowledge (ECMA-376 Part 1, MS-OI29500): a 15-bit rolling XOR over
 * the password's character codes in reverse order, finalized with the length
 * and the magic constant 0xCE4B. Excel writes the result as uppercase hex,
 * zero-padded to 4 digits — the same encoding openpyxl reads back into
 * `ws.protection.password`.
 */
export function excelLegacyPasswordHash(password: string): string {
  let hash = 0
  for (let index = password.length - 1; index >= 0; index -= 1) {
    hash = (((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)) ^ password.charCodeAt(index)
  }
  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff)
  hash ^= password.length
  hash ^= 0xce4b
  return hash.toString(16).toUpperCase().padStart(4, '0')
}

/// Case-insensitive compare of a candidate against a stored legacy hash
/// (Excel treats the attribute as hex, so case never matters).
export function legacyPasswordMatches(password: string, storedHash: string): boolean {
  return excelLegacyPasswordHash(password) === storedHash.toUpperCase()
}
