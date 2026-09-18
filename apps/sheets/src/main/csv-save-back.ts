import { atomicWriteFile } from '@airy-office/electron-utils'

/** UTF-8 BOM so Excel decodes the reopened file correctly. */
export function csvBytesWithBom(content: string): Buffer {
  return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(content, 'utf8')])
}

/**
 * A CSV in-place save writes back into the user's original .csv — usually the
 * only copy. A plain writeFile truncates the target the moment it opens it,
 * so a crash or disk error mid-write destroys the file; stage the bytes in a
 * sibling temp file and rename them into place instead, the same guarantee
 * the xlsx save path already has.
 */
export async function writeCsvBackAtomic(path: string, content: string): Promise<void> {
  await atomicWriteFile(path, csvBytesWithBom(content))
}
