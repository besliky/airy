// OLE2/CFB container sniffing for the open paths (BUG-1504): an encrypted
// Office document is not a zip — Word repackages it as a Compound File Binary
// (OLE2) container holding an EncryptedPackage stream (encrypted OOXML:
// .docx/.xlsx/.pptx with a password) or sets the FIB fEncrypted flag (legacy
// encrypted .doc). Without this sniff a password-protected .docx surfaced as
// jszip's "Can't find end of central directory" (looks like a corrupt file)
// and an encrypted .doc as word-extractor's "Attempt to access memory outside
// buffer bounds" — neither tells the agent the file is password-protected.
// The same detection exists in the desktop apps (apps/docs
// /src/main/docx-encryption.ts, apps/slides/src/main/cfb-sniff.ts); the
// headless server vendors the *refusal* only: decryption needs an interactive
// password prompt, which an MCP session cannot provide.
import { open } from 'node:fs/promises'

const CFB_MAGIC = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
// stream name as stored in a CFB directory (UTF-16LE) — marks encrypted OOXML
const ENCRYPTED_PACKAGE_UTF16LE = Uint8Array.from(
  'EncryptedPackage'.split('').flatMap((ch) => [ch.charCodeAt(0), 0]),
)
/** FIB base flags (uint16 at offset 0x0A): bit 0x0100 = fEncrypted */
const FIB_FLAGS_OFFSET = 0x0a
const FIB_ENCRYPTED_FLAG = 0x0100

/** how much of a file the conversion-path head sniff reads (the fixtures'
 *  EncryptedPackage directories sit in the first KiB; a full scan happens on
 *  the docx session path, which already holds every byte) */
export const OLE_SNIFF_MAX_BYTES = 64 * 1024

/** typed refusal so the text-extraction open rethrows it unwrapped */
export class EncryptedOfficeError extends Error {
  constructor(path: string, format: string) {
    super(
      `Cannot open "${path}": the file is password-protected (an encrypted Office container, ` +
        `not plain .${format} bytes). Headless sessions cannot decrypt it: remove the password ` +
        'in the source app (e.g. Word: File > Info > Protect Document) or save an unprotected ' +
        'copy, then open that.',
    )
    this.name = 'EncryptedOfficeError'
  }
}

export type OleContent = 'encrypted-ooxml' | 'encrypted-legacy' | 'plain'

/** true when the bytes start with the Compound File Binary (OLE2) magic */
export function isOleContainer(bytes: Uint8Array): boolean {
  return bytes.length >= CFB_MAGIC.length && CFB_MAGIC.every((byte, index) => bytes[index] === byte)
}

/**
 * Classify CFB content: encrypted OOXML (an EncryptedPackage stream — the
 * UTF-16LE stream name appears in the directory), a legacy encrypted .doc
 * (FIB fEncrypted flag), or a plain OLE document (e.g. an unencrypted .doc).
 * Non-CFB bytes return null. The needle search runs on a Buffer view so the
 * scan is memmem-fast even for a maximum-size hostile container.
 */
export function classifyOleContent(bytes: Uint8Array): OleContent | null {
  if (!isOleContainer(bytes)) return null
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const needle = Buffer.from(
    ENCRYPTED_PACKAGE_UTF16LE.buffer,
    ENCRYPTED_PACKAGE_UTF16LE.byteOffset,
    ENCRYPTED_PACKAGE_UTF16LE.byteLength,
  )
  if (view.includes(needle)) return 'encrypted-ooxml'
  if (bytes.length >= FIB_FLAGS_OFFSET + 2) {
    const flags = bytes[FIB_FLAGS_OFFSET]! | (bytes[FIB_FLAGS_OFFSET + 1]! << 8)
    if ((flags & FIB_ENCRYPTED_FLAG) !== 0) return 'encrypted-legacy'
  }
  return 'plain'
}

/** the encrypted-file refusal for an open path (typed; see EncryptedOfficeError) */
export function encryptedOfficeRefusal(path: string, format: string): EncryptedOfficeError {
  return new EncryptedOfficeError(path, format)
}

/** Read at most `max` bytes from the start of a file (head sniffing). */
export async function readHead(path: string, max: number): Promise<Uint8Array> {
  const handle = await open(path, 'r')
  try {
    const head = new Uint8Array(max)
    const { bytesRead } = await handle.read(head, 0, max, 0)
    return head.subarray(0, bytesRead)
  } finally {
    await handle.close()
  }
}
