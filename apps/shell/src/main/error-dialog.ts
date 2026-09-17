import { dialog } from 'electron'
import type { BrowserWindow } from 'electron'
import { createI18n, getUiLang } from '@airy-office/i18n'
import { friendlyErrorKey } from '../shared/error-codes'

let showing = false

/** friendly text per mapped filesystem error code; localized here so the
 *  dialog stays self-contained (no import cycle with the shell dictionary) */
const tFriendly = createI18n({
  zh: {
    enoent: '文件不存在或已被移动。',
    eperm: '没有访问该文件的权限。',
    ebusy: '文件正被其他程序占用，请关闭后重试。',
    emfile: '打开的文件过多，请稍后重试。',
  },
  en: {
    enoent: 'The file does not exist or has been moved.',
    eperm: 'Permission denied for this file.',
    ebusy: 'The file is open in another program. Close it and try again.',
    emfile: 'Too many files are open. Try again shortly.',
  },
  ja: {
    enoent: 'ファイルが存在しないか、移動されました。',
    eperm: 'このファイルへのアクセスが拒否されました。',
    ebusy: 'ファイルが他のプログラムで開いています。閉じてから再試行してください。',
    emfile: '開いているファイルが多すぎます。しばらくしてから再試行してください。',
  },
  ko: {
    enoent: '파일이 없거나 이동되었습니다.',
    eperm: '이 파일에 대한 액세스 권한이 없습니다.',
    ebusy: '파일이 다른 프로그램에서 열려 있습니다. 닫고 다시 시도하세요.',
    emfile: '열려 있는 파일이 너무 많습니다. 잠시 후 다시 시도하세요.',
  },
  fr: {
    enoent: "Le fichier n'existe pas ou a été déplacé.",
    eperm: 'Accès refusé à ce fichier.',
    ebusy: 'Le fichier est ouvert dans un autre programme. Fermez-le puis réessayez.',
    emfile: 'Trop de fichiers sont ouverts. Réessayez dans un instant.',
  },
  de: {
    enoent: 'Die Datei ist nicht vorhanden oder wurde verschoben.',
    eperm: 'Zugriff auf diese Datei verweigert.',
    ebusy:
      'Die Datei ist in einem anderen Programm geöffnet. Schließen Sie es und versuchen Sie es erneut.',
    emfile: 'Zu viele Dateien sind geöffnet. Versuchen Sie es bald erneut.',
  },
  es: {
    enoent: 'El archivo no existe o se ha movido.',
    eperm: 'Permiso denegado para este archivo.',
    ebusy: 'El archivo está abierto en otro programa. Ciérralo e inténtalo de nuevo.',
    emfile: 'Hay demasiados archivos abiertos. Inténtalo de nuevo en un momento.',
  },
  th: {
    enoent: 'ไฟล์ไม่มีอยู่หรือถูกย้ายไปแล้ว',
    eperm: 'ไม่มีสิทธิ์เข้าถึงไฟล์นี้',
    ebusy: 'ไฟล์ถูกเปิดในโปรแกรมอื่นอยู่ ปิดแล้วลองอีกครั้ง',
    emfile: 'เปิดไฟล์มากเกินไป โปรดลองอีกครั้งในอีกครู่',
  },
  id: {
    enoent: 'Berkas tidak ada atau telah dipindahkan.',
    eperm: 'Akses ke berkas ini ditolak.',
    ebusy: 'Berkas sedang dibuka di program lain. Tutup lalu coba lagi.',
    emfile: 'Terlalu banyak berkas terbuka. Coba lagi sebentar lagi.',
  },
  ru: {
    enoent: 'Файл не существует или был перемещён.',
    eperm: 'Доступ к файлу запрещён.',
    ebusy: 'Файл открыт в другой программе. Закройте её и повторите попытку.',
    emfile: 'Открыто слишком много файлов. Повторите попытку позже.',
  },
  ar: {
    enoent: 'الملف غير موجود أو تم نقله.',
    eperm: 'تم رفض الوصول إلى هذا الملف.',
    ebusy: 'الملف مفتوح في برنامج آخر. أغلقه ثم أعد المحاولة.',
    emfile: 'عدد الملفات المفتوحة كبير جدًا. أعد المحاولة بعد قليل.',
  },
  pt: {
    enoent: 'O arquivo não existe ou foi movido.',
    eperm: 'Acesso negado a este arquivo.',
    ebusy: 'O arquivo está aberto em outro programa. Feche-o e tente novamente.',
    emfile: 'Há muitos arquivos abertos. Tente novamente em instantes.',
  },
  it: {
    enoent: 'Il file non esiste o è stato spostato.',
    eperm: 'Accesso negato a questo file.',
    ebusy: 'Il file è aperto in un altro programma. Chiudilo e riprova.',
    emfile: 'Troppi file aperti. Riprova tra poco.',
  },
  pl: {
    enoent: 'Plik nie istnieje lub został przeniesiony.',
    eperm: 'Brak dostępu do tego pliku.',
    ebusy: 'Plik jest otwarty w innym programie. Zamknij go i spróbuj ponownie.',
    emfile: 'Otwarto zbyt wiele plików. Spróbuj ponownie za chwilę.',
  },
  cs: {
    enoent: 'Soubor neexistuje nebo byl přesunut.',
    eperm: 'Přístup k tomuto souboru byl odepřen.',
    ebusy: 'Soubor je otevřený v jiném programu. Zavřete ho a zkuste to znovu.',
    emfile: 'Je otevřeno příliš mnoho souborů. Zkuste to za chvíli znovu.',
  },
  nl: {
    enoent: 'Het bestand bestaat niet of is verplaatst.',
    eperm: 'Toegang tot dit bestand geweigerd.',
    ebusy: 'Het bestand is geopend in een ander programma. Sluit het en probeer opnieuw.',
    emfile: 'Er zijn te veel bestanden open. Probeer het zo opnieuw.',
  },
  ms: {
    enoent: 'Fail tidak wujud atau telah dipindahkan.',
    eperm: 'Akses kepada fail ini ditolak.',
    ebusy: 'Fail dibuka dalam program lain. Tutup dan cuba lagi.',
    emfile: 'Terlalu banyak fail dibuka. Cuba lagi sebentar lagi.',
  },
  he: {
    enoent: 'הקובץ אינו קיים או הועבר.',
    eperm: 'הגישה לקובץ זה נדחתה.',
    ebusy: 'הקובץ פתוח בתוכנית אחרת. סגרו אותה ונסו שוב.',
    emfile: 'יותר מדי קבצים פתוחים. נסו שוב בעוד רגע.',
  },
  hi: {
    enoent: 'फ़ाइल मौजूद नहीं है या ले जाई गई है।',
    eperm: 'इस फ़ाइल तक पहुँच अस्वीकृत है।',
    ebusy: 'फ़ाइल किसी अन्य प्रोग्राम में खुली है। बंद करके पुनः प्रयास करें।',
    emfile: 'बहुत अधिक फ़ाइलें खुली हैं। थोड़ी देर बाद पुनः प्रयास करें।',
  },
  'zh-TW': {
    enoent: '檔案不存在或已被移動。',
    eperm: '沒有存取此檔案的權限।',
    ebusy: '檔案正被其他程式佔用，請關閉後再試一次。',
    emfile: '開啟的檔案過多，請稍後再試一次。',
  },
})

/**
 * Never dialog.showErrorBox here: on Windows it blocks main-process JS in a
 * nested native pump and, parentless, can hide behind the window — a wedged
 * UI with no "(Not Responding)". Async + parented avoids both.
 *
 * Known filesystem errors (ENOENT, EACCES/EPERM, EBUSY/ETXTBSY, EMFILE) are
 * prefixed with a friendly localized explanation; the raw error text stays
 * visible as the dialog's detail section.
 */
export function showErrorDialog(win: BrowserWindow | null, message: string, err: unknown): void {
  if (showing) return
  showing = true
  const raw = err instanceof Error ? err.message : String(err)
  const friendly = friendlyErrorKey(err)
  const options = {
    type: 'error' as const,
    message,
    detail: friendly ? `${tFriendly(getUiLang(), friendly)}\n\n${raw}` : raw,
  }
  const shown =
    win && !win.isDestroyed() ? dialog.showMessageBox(win, options) : dialog.showMessageBox(options)
  void shown.finally(() => {
    showing = false
  })
}
