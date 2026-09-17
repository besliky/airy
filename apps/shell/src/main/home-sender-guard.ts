/**
 * Sender validation for the process-global home:* IPC channels.
 *
 * The shell bundles every editor's main code into one process, so an
 * ipcMain.handle registered for the Home screen answers any webContents —
 * a compromised editor renderer could otherwise rename/trash/duplicate
 * files through them. The only legitimate sender is the Home tab, which is
 * the shell window's own renderer (its webContents id is captured at window
 * creation). Pure so the decision is unit-testable.
 */
export function isHomeSender(
  homeWebContentsId: number | null | undefined,
  senderId: number | undefined,
): boolean {
  return (
    homeWebContentsId !== null && homeWebContentsId !== undefined && senderId === homeWebContentsId
  )
}
