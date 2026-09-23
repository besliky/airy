import { existsSync, mkdirSync } from 'node:fs'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, join, relative, resolve, sep } from 'node:path'
import { pathToFileURL } from 'node:url'
import {
  BrowserWindow,
  WebContentsView,
  app,
  dialog,
  ipcMain,
  net,
  protocol,
  shell,
} from 'electron'
import type { WebContents } from 'electron'
import {
  TextRecoveryStore,
  checkSaveStaleness,
  configuredDefaultSaveDir,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  statFileStamp,
  truncateByCodePoints,
  voidLoad,
} from '@airy-office/electron-utils'
import type { FileStamp } from '@airy-office/electron-utils'
import { createI18n, getUiLang } from '@airy-office/i18n'
import type { Params } from '@airy-office/i18n'
import { generateImageTool } from '@airy-office/ai-search'
import { decodeTextBytes, legacyCharsetForLang } from '@airy-office/file-parse/text'
import { atomicWriteFile } from '@airy-office/electron-utils'
import {
  copyImageIntoOwnedAssets,
  discardPendingOwnedAssets,
  extractMarkdownImageSources,
  pendingOwnedAssetsForDocument,
  prepareAssetsForSaveAs,
  reconcileOwnedAssets,
  renameOwnedAssetDocument,
  resolveSafeRelativeImagePath,
  resolveSourcePendingAfterSaveAs,
  rollbackPreparedSaveAsAssets,
  writeImageIntoOwnedAssets,
} from './asset-lifecycle'
import { createMarkdownConversionSession, writeMarkdownConversion } from './conversion-lifecycle'
import { MARKDOWN_CHANNELS } from '../shared/ipc'
import type {
  ExportDocxRequest,
  ExportFormat,
  ExportPdfRequest,
  ExportResult,
  ImageData,
  SaveMarkdownRequest,
  SaveMarkdownResult,
  SaveMode,
} from '../shared/ipc'

const tDlg = createI18n({
  zh: {
    dlgSaveTitle: '保存 Markdown 文档',
    filterMarkdown: 'Markdown 文档',
    dlgPickImage: '选择图片',
    filterImages: '图片',
    untitledFile: '未命名文档',
    externalChangeTitle: '文件已被外部修改',
    externalChangeDetail:
      '自上次保存后，“{name}”已被其他程序更改、重命名或删除。覆盖将丢弃外部更改；另存为会将您的版本写入新文件。',
    btnOverwrite: '覆盖',
    btnSaveAs: '另存为',
    closeUnsavedMsg: '此文档有未保存的更改。',
    autosaveFoundTitle: '发现自动恢复版本',
    autosaveFoundBody: '上次会话有未保存的更改。要恢复自动保存的版本吗?',
    autosaveRestore: '恢复',
    autosaveDiscard: '放弃',
    closeUnsavedDetail: '关闭前是否保存？',
    btnSave: '保存',
    btnDontSave: '不保存',
    btnCancel: '取消',
  },
  en: {
    dlgSaveTitle: 'Save Markdown Document',
    filterMarkdown: 'Markdown Documents',
    dlgPickImage: 'Choose an Image',
    filterImages: 'Images',
    untitledFile: 'Untitled',
    externalChangeTitle: 'File changed outside this app',
    externalChangeDetail:
      '"{name}" was changed, renamed, or deleted by another program since the last save. Overwriting discards the external changes; Save As writes your version to a new file.',
    btnOverwrite: 'Overwrite',
    btnSaveAs: 'Save As',
    closeUnsavedMsg: 'This document has unsaved changes.',
    autosaveFoundTitle: 'Recovered version found',
    autosaveFoundBody:
      'There are unsaved changes from your last session. Restore the autosaved version?',
    autosaveRestore: 'Restore',
    autosaveDiscard: 'Discard',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    btnSave: 'Save',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
  },
  ja: {
    dlgSaveTitle: 'Markdown ドキュメントを保存',
    filterMarkdown: 'Markdown ドキュメント',
    dlgPickImage: '画像を選択',
    filterImages: '画像',
    untitledFile: '無題',
    externalChangeTitle: 'ファイルが外部で変更されました',
    externalChangeDetail:
      '最後の保存後、「{name}」は他のプログラムによって変更・名前変更・削除されました。上書きすると外部の変更は失われます。名前を付けて保存すると新しいファイルに保存されます。',
    btnOverwrite: '上書き保存',
    btnSaveAs: '名前を付けて保存',
    closeUnsavedMsg: 'このドキュメントに未保存の変更があります。',
    autosaveFoundTitle: '自動回復バージョンがあります',
    autosaveFoundBody: '前回のセッションに未保存の変更があります。自動保存版を復元しますか?',
    autosaveRestore: '復元',
    autosaveDiscard: '破棄',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    btnSave: '保存',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
  },
  ko: {
    dlgSaveTitle: 'Markdown 문서 저장',
    filterMarkdown: 'Markdown 문서',
    dlgPickImage: '이미지 선택',
    filterImages: '이미지',
    untitledFile: '제목 없음',
    externalChangeTitle: '파일이 외부에서 변경되었습니다',
    externalChangeDetail:
      '마지막 저장 후 "{name}"이(가) 다른 프로그램에 의해 변경, 이름 변경 또는 삭제되었습니다. 덮어쓰면 외부 변경 사항이 손실됩니다. 다른 이름으로 저장하면 새 파일에 저장됩니다.',
    btnOverwrite: '덮어쓰기',
    btnSaveAs: '다른 이름으로 저장',
    closeUnsavedMsg: '이 문서에 저장하지 않은 변경 사항이 있습니다.',
    autosaveFoundTitle: '자동 복구 버전 발견',
    autosaveFoundBody:
      '마지막 세션에 저장되지 않은 변경 내용이 있습니다. 자동 저장 버전을 복원할까요?',
    autosaveRestore: '복원',
    autosaveDiscard: '취소',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    btnSave: '저장',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
  },
  fr: {
    dlgSaveTitle: 'Enregistrer le document Markdown',
    filterMarkdown: 'Documents Markdown',
    dlgPickImage: 'Choisir une image',
    filterImages: 'Images',
    untitledFile: 'Sans titre',
    externalChangeTitle: 'Fichier modifié à l’extérieur de l’application',
    externalChangeDetail:
      'Depuis le dernier enregistrement, « {name} » a été modifié, renommé ou supprimé par un autre programme. Écraser abandonne les modifications externes ; Enregistrer sous écrit votre version dans un nouveau fichier.',
    btnOverwrite: 'Écraser',
    btnSaveAs: 'Enregistrer sous',
    closeUnsavedMsg: 'Ce document contient des modifications non enregistrées.',
    autosaveFoundTitle: 'Version récupérée trouvée',
    autosaveFoundBody:
      'Des modifications non enregistrées existent. Restaurer la version auto-enregistrée ?',
    autosaveRestore: 'Restaurer',
    autosaveDiscard: 'Ignorer',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    btnSave: 'Enregistrer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
  },
  de: {
    dlgSaveTitle: 'Markdown-Dokument speichern',
    filterMarkdown: 'Markdown-Dokumente',
    dlgPickImage: 'Bild auswählen',
    filterImages: 'Bilder',
    untitledFile: 'Unbenannt',
    externalChangeTitle: 'Datei außerhalb der App geändert',
    externalChangeDetail:
      '„{name}“ wurde seit dem letzten Speichern von einem anderen Programm geändert, umbenannt oder gelöscht. Überschreiben verwirft die externen Änderungen; Speichern unter schreibt Ihre Version in eine neue Datei.',
    btnOverwrite: 'Überschreiben',
    btnSaveAs: 'Speichern unter',
    closeUnsavedMsg: 'Dieses Dokument enthält ungespeicherte Änderungen.',
    autosaveFoundTitle: 'Wiederhergestellte Version gefunden',
    autosaveFoundBody:
      'Es gibt ungespeicherte Änderungen. Automatisch gespeicherte Version wiederherstellen?',
    autosaveRestore: 'Wiederherstellen',
    autosaveDiscard: 'Verwerfen',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    btnSave: 'Speichern',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
  },
  es: {
    dlgSaveTitle: 'Guardar documento Markdown',
    filterMarkdown: 'Documentos Markdown',
    dlgPickImage: 'Elegir imagen',
    filterImages: 'Imágenes',
    untitledFile: 'Sin título',
    externalChangeTitle: 'Archivo modificado fuera de la aplicación',
    externalChangeDetail:
      '«{name}» fue cambiado, renombrado o eliminado por otro programa desde el último guardado. Sobrescribir descarta los cambios externos; Guardar como escribe su versión en un archivo nuevo.',
    btnOverwrite: 'Sobrescribir',
    btnSaveAs: 'Guardar como',
    closeUnsavedMsg: 'Este documento tiene cambios sin guardar.',
    autosaveFoundTitle: 'Se encontró una versión recuperada',
    autosaveFoundBody:
      'Hay cambios sin guardar de la última sesión. ¿Restaurar la versión autoguardada?',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    closeUnsavedDetail: '¿Quieres guardarlos antes de cerrar?',
    btnSave: 'Guardar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
  },
  th: {
    dlgSaveTitle: 'บันทึกเอกสาร Markdown',
    filterMarkdown: 'เอกสาร Markdown',
    dlgPickImage: 'เลือกรูปภาพ',
    filterImages: 'รูปภาพ',
    untitledFile: 'ไม่มีชื่อ',
    externalChangeTitle: 'ไฟล์ถูกแก้ไขจากภายนอกแอป',
    externalChangeDetail:
      '"{name}" ถูกเปลี่ยน เปลี่ยนชื่อ หรือลบโดยโปรแกรมอื่นตั้งแต่การบันทึกครั้งล่าสุด การเขียนทับจะละทิ้งการเปลี่ยนแปลงภายนอก ส่วนบันทึกเป็นจะเขียนเวอร์ชันของคุณลงไฟล์ใหม่',
    btnOverwrite: 'เขียนทับ',
    btnSaveAs: 'บันทึกเป็น',
    closeUnsavedMsg: 'เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    autosaveFoundTitle: 'พบเวอร์ชันกู้คืนอัตโนมัติ',
    autosaveFoundBody: 'มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึกจากครั้งก่อน ต้องการกู้คืนหรือไม่?',
    autosaveRestore: 'กู้คืน',
    autosaveDiscard: 'ละทิ้ง',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    btnSave: 'บันทึก',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
  },
  id: {
    dlgSaveTitle: 'Simpan dokumen Markdown',
    filterMarkdown: 'Dokumen Markdown',
    dlgPickImage: 'Pilih gambar',
    filterImages: 'Gambar',
    untitledFile: 'Tanpa judul',
    externalChangeTitle: 'File diubah dari luar aplikasi',
    externalChangeDetail:
      '"{name}" diubah, diganti nama, atau dihapus oleh program lain sejak penyimpanan terakhir. Timpa akan membuang perubahan eksternal; Simpan Sebagai menulis versi Anda ke file baru.',
    btnOverwrite: 'Timpa',
    btnSaveAs: 'Simpan Sebagai',
    closeUnsavedMsg: 'Dokumen ini memiliki perubahan yang belum disimpan.',
    autosaveFoundTitle: 'Versi pemulihan ditemukan',
    autosaveFoundBody:
      'Ada perubahan yang belum disimpan dari sesi terakhir. Pulihkan versi tersimpan otomatis?',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  ru: {
    dlgSaveTitle: 'Сохранить документ Markdown',
    filterMarkdown: 'Документы Markdown',
    dlgPickImage: 'Выберите изображение',
    filterImages: 'Изображения',
    untitledFile: 'Без названия',
    externalChangeTitle: 'Файл изменён вне приложения',
    externalChangeDetail:
      'С момента последнего сохранения «{name}» был изменён, переименован или удалён другой программой. Перезапись отменит внешние изменения; «Сохранить как» запишет вашу версию в новый файл.',
    btnOverwrite: 'Перезаписать',
    btnSaveAs: 'Сохранить как',
    closeUnsavedMsg: 'В этом документе есть несохранённые изменения.',
    autosaveFoundTitle: 'Найдена восстановленная версия',
    autosaveFoundBody:
      'Есть несохранённые изменения из прошлого сеанса. Восстановить автосохранённую версию?',
    autosaveRestore: 'Восстановить',
    autosaveDiscard: 'Отклонить',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    btnSave: 'Сохранить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
  },
  ar: {
    dlgSaveTitle: 'حفظ مستند Markdown',
    filterMarkdown: 'مستندات Markdown',
    dlgPickImage: 'اختر صورة',
    filterImages: 'صور',
    untitledFile: 'بدون عنوان',
    externalChangeTitle: 'تم تغيير الملف من خارج التطبيق',
    externalChangeDetail:
      'تم تغيير "{name}" أو إعادة تسميته أو حذفه بواسطة برنامج آخر منذ آخر حفظ. الكتابة فوق تتخلى عن التغييرات الخارجية؛ حفظ باسم يكتب نسختك في ملف جديد.',
    btnOverwrite: 'الكتابة فوق',
    btnSaveAs: 'حفظ باسم',
    closeUnsavedMsg: 'يحتوي هذا المستند على تغييرات غير محفوظة.',
    autosaveFoundTitle: 'تم العثور على نسخة مستردة',
    autosaveFoundBody:
      'توجد تغييرات غير محفوظة من الجلسة الأخيرة. هل تريد استعادة النسخة المحفوظة تلقائيًا؟',
    autosaveRestore: 'استعادة',
    autosaveDiscard: 'تجاهل',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    btnSave: 'حفظ',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
  },
  pt: {
    dlgSaveTitle: 'Salvar documento Markdown',
    filterMarkdown: 'Documentos Markdown',
    dlgPickImage: 'Escolher imagem',
    filterImages: 'Imagens',
    untitledFile: 'Sem título',
    externalChangeTitle: 'Arquivo alterado fora do aplicativo',
    externalChangeDetail:
      '"{name}" foi alterado, renomeado ou excluído por outro programa desde o último salvamento. Sobrescrever descarta as alterações externas; Salvar como grava a sua versão em um novo arquivo.',
    btnOverwrite: 'Sobrescrever',
    btnSaveAs: 'Salvar como',
    closeUnsavedMsg: 'Este documento tem alterações não salvas.',
    autosaveFoundTitle: 'Versão recuperada encontrada',
    autosaveFoundBody:
      'Há alterações não salvas da sua última sessão. Restaurar a versão salva automaticamente?',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    btnSave: 'Salvar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
  },
  it: {
    dlgSaveTitle: 'Salva documento Markdown',
    filterMarkdown: 'Documenti Markdown',
    dlgPickImage: 'Scegli immagine',
    filterImages: 'Immagini',
    untitledFile: 'Senza titolo',
    externalChangeTitle: 'File modificato al di fuori dell’app',
    externalChangeDetail:
      '"{name}" è stato modificato, rinominato o eliminato da un altro programma dall’ultimo salvataggio. Sovrascrivere scarta le modifiche esterne; Salva come scrive la tua versione in un nuovo file.',
    btnOverwrite: 'Sovrascrivi',
    btnSaveAs: 'Salva come',
    closeUnsavedMsg: 'Questo documento contiene modifiche non salvate.',
    autosaveFoundTitle: 'Trovata versione recuperata',
    autosaveFoundBody:
      "Ci sono modifiche non salvate dall'ultima sessione. Ripristinare la versione salvata automaticamente?",
    autosaveRestore: 'Ripristina',
    autosaveDiscard: 'Ignora',
    closeUnsavedDetail: 'Vuoi salvarle prima di chiudere?',
    btnSave: 'Salva',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
  },
  pl: {
    dlgSaveTitle: 'Zapisz dokument Markdown',
    filterMarkdown: 'Dokumenty Markdown',
    dlgPickImage: 'Wybierz obraz',
    filterImages: 'Obrazy',
    untitledFile: 'Bez tytułu',
    externalChangeTitle: 'Plik zmieniony poza aplikacją',
    externalChangeDetail:
      'Od ostatniego zapisu plik „{name}” został zmieniony, przemianowany lub usunięty przez inny program. Nadpisanie odrzuci zmiany zewnętrzne; Zapisz jako zapisze Twoją wersję w nowym pliku.',
    btnOverwrite: 'Nadpisz',
    btnSaveAs: 'Zapisz jako',
    closeUnsavedMsg: 'Ten dokument ma niezapisane zmiany.',
    autosaveFoundTitle: 'Znaleziono odzyskaną wersję',
    autosaveFoundBody:
      'Istnieją niezapisane zmiany z ostatniej sesji. Przywrócić wersję zapisaną automatycznie?',
    autosaveRestore: 'Przywróć',
    autosaveDiscard: 'Odrzuć',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    btnSave: 'Zapisz',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
  },
  cs: {
    dlgSaveTitle: 'Uložit dokument Markdown',
    filterMarkdown: 'Dokumenty Markdown',
    dlgPickImage: 'Vyberte obrázek',
    filterImages: 'Obrázky',
    untitledFile: 'Bez názvu',
    externalChangeTitle: 'Soubor byl změněn mimo aplikaci',
    externalChangeDetail:
      'Od posledního uložení byl soubor „{name}“ změněn, přejmenován nebo smazán jiným programem. Přepsáním se zahodí externí změny; Uložit jako zapíše vaši verzi do nového souboru.',
    btnOverwrite: 'Přepsat',
    btnSaveAs: 'Uložit jako',
    closeUnsavedMsg: 'Tento dokument má neuložené změny.',
    autosaveFoundTitle: 'Nalezena obnovená verze',
    autosaveFoundBody:
      'Z poslední relace existují neuložené změny. Obnovit automaticky uloženou verzi?',
    autosaveRestore: 'Obnovit',
    autosaveDiscard: 'Zahodit',
    closeUnsavedDetail: 'Chcete je před zavřením uložit?',
    btnSave: 'Uložit',
    btnDontSave: 'Neukládat',
    btnCancel: 'Zrušit',
  },
  nl: {
    dlgSaveTitle: 'Markdown-document opslaan',
    filterMarkdown: 'Markdown-documenten',
    dlgPickImage: 'Kies een afbeelding',
    filterImages: 'Afbeeldingen',
    untitledFile: 'Naamloos',
    externalChangeTitle: 'Bestand buiten de app gewijzigd',
    externalChangeDetail:
      '"{name}" is sinds de laatste opslag gewijzigd, hernoemd of verwijderd door een ander programma. Overschrijven staakt de externe wijzigingen; Opslaan als schrijft uw versie naar een nieuw bestand.',
    btnOverwrite: 'Overschrijven',
    btnSaveAs: 'Opslaan als',
    closeUnsavedMsg: 'Dit document bevat niet-opgeslagen wijzigingen.',
    autosaveFoundTitle: 'Herstelde versie gevonden',
    autosaveFoundBody:
      'Er zijn niet-opgeslagen wijzigingen van uw laatste sessie. De automatisch opgeslagen versie herstellen?',
    autosaveRestore: 'Herstellen',
    autosaveDiscard: 'Negeren',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    btnSave: 'Opslaan',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
  },
  ms: {
    dlgSaveTitle: 'Simpan dokumen Markdown',
    filterMarkdown: 'Dokumen Markdown',
    dlgPickImage: 'Pilih imej',
    filterImages: 'Imej',
    untitledFile: 'Tanpa tajuk',
    externalChangeTitle: 'Fail diubah di luar aplikasi',
    externalChangeDetail:
      '"{name}" telah diubah, dinamakan semula atau dipadamkan oleh program lain sejak simpanan terakhir. Tulis ganti akan mengetepikan perubahan luaran; Simpan Sebagai menulis versi anda ke fail baharu.',
    btnOverwrite: 'Tulis ganti',
    btnSaveAs: 'Simpan Sebagai',
    closeUnsavedMsg: 'Dokumen ini mempunyai perubahan yang belum disimpan.',
    autosaveFoundTitle: 'Versi pulihan ditemui',
    autosaveFoundBody:
      'Terdapat perubahan yang belum disimpan daripada sesi terakhir anda. Pulihkan versi yang disimpan secara automatik?',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  he: {
    dlgSaveTitle: 'שמירת מסמך Markdown',
    filterMarkdown: 'מסמכי Markdown',
    dlgPickImage: 'בחרו תמונה',
    filterImages: 'תמונות',
    untitledFile: 'ללא שם',
    externalChangeTitle: 'הקובץ השתנה מחוץ ליישום',
    externalChangeDetail:
      '"{name}" שונה, שונה שמו או נמחק על ידי תוכנית אחרת מאז השמירה האחרונה. דריסה מוותרת על השינויים החיצוניים; שמירה בשם כותבת את הגרסה שלך לקובץ חדש.',
    btnOverwrite: 'דריסה',
    btnSaveAs: 'שמירה בשם',
    closeUnsavedMsg: 'במסמך הזה יש שינויים שלא נשמרו.',
    autosaveFoundTitle: 'נמצאה גרסה משוחזרת',
    autosaveFoundBody: 'קיימים שינויים שלא נשמרו מהפעלה הקודמת. לשחזר את הגרסה שנשמרה אוטומטית?',
    autosaveRestore: 'שחזר',
    autosaveDiscard: 'התעלם',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    btnSave: 'שמירה',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
  },
  hi: {
    dlgSaveTitle: 'Markdown दस्तावेज़ सहेजें',
    filterMarkdown: 'Markdown दस्तावेज़',
    dlgPickImage: 'छवि चुनें',
    filterImages: 'छवियाँ',
    untitledFile: 'शीर्षकहीन',
    externalChangeTitle: 'फ़ाइल ऐप के बाहर बदली गई',
    externalChangeDetail:
      'पिछले सहेजने के बाद "{name}" को किसी अन्य प्रोग्राम ने बदला, नाम बदला या हटाया। ओवरराइट करने पर बाहरी बदलाव खो जाएंगे; इस रूप में सहेजें आपका संस्करण नई फ़ाइल में लिखेगा।',
    btnOverwrite: 'ओवरराइट करें',
    btnSaveAs: 'इस रूप में सहेजें',
    closeUnsavedMsg: 'इस दस्तावेज़ में सहेजे नहीं गए परिवर्तन हैं।',
    autosaveFoundTitle: 'पुनर्प्राप्त संस्करण मिला',
    autosaveFoundBody:
      'आपके पिछले सत्र से सहेजे नहीं गए परिवर्तन हैं। स्वतः सहेजा गया संस्करण पुनर्स्थापित करें?',
    autosaveRestore: 'पुनर्स्थापित करें',
    autosaveDiscard: 'छोड़ें',
    closeUnsavedDetail: 'क्या बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    btnSave: 'सहेजें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
  },
  'zh-TW': {
    dlgSaveTitle: '儲存 Markdown 文件',
    filterMarkdown: 'Markdown 文件',
    dlgPickImage: '選擇圖片',
    filterImages: '圖片',
    untitledFile: '未命名文件',
    externalChangeTitle: '檔案已被外部修改',
    externalChangeDetail:
      '自上次儲存後，「{name}」已被其他程式變更、重新命名或刪除。覆蓋將捨棄外部變更；另存新檔會將您的版本寫入新檔案。',
    btnOverwrite: '覆蓋',
    btnSaveAs: '另存新檔',
    closeUnsavedMsg: '此文件有未儲存的變更。',
    autosaveFoundTitle: '發現自動復原版本',
    autosaveFoundBody: '上次工作階段有未儲存的變更。要復原自動儲存的版本嗎?',
    autosaveRestore: '復原',
    autosaveDiscard: '放棄',
    closeUnsavedDetail: '關閉前是否儲存？',
    btnSave: '儲存',
    btnDontSave: '不儲存',
    btnCancel: '取消',
  },
})
type DlgKey =
  | 'dlgSaveTitle'
  | 'filterMarkdown'
  | 'dlgPickImage'
  | 'filterImages'
  | 'untitledFile'
  | 'closeUnsavedMsg'
  | 'closeUnsavedDetail'
  | 'btnSave'
  | 'btnDontSave'
  | 'btnCancel'
  | 'autosaveFoundTitle'
  | 'autosaveFoundBody'
  | 'autosaveRestore'
  | 'autosaveDiscard'
  | 'externalChangeTitle'
  | 'externalChangeDetail'
  | 'btnOverwrite'
  | 'btnSaveAs'
const tm = (key: DlgKey, params?: Params) => tDlg(getUiLang(), key, params)

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
  /** Shell router used to open exported PDFs in a new Airy tab; the asking
   *  view's webContents id rides along so the tab opens in the sender's
   *  window, not whichever window holds focus (BUG-1107 focus routing). */
  openGeneratedPath?: (path: string, senderWcId?: number) => boolean
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configureMarkdownRuntime(paths: RuntimePaths): void {
  runtime = paths
}

/** After a successful Markdown → PDF export: open the file in a PDF tab (shell)
 * or reveal it in the folder (standalone). Tab-opening failure must not
 * report the export itself as failed — the file is already persisted. */
function openExportedPdf(path: string, senderWcId?: number): void {
  try {
    if (runtime.openGeneratedPath?.(path, senderWcId)) return
  } catch (err) {
    console.warn('[markdown] Failed to open exported PDF:', err)
  }
  shell.showItemInFolder(path)
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount.
 * Kept until the view is destroyed so a reload (View > Reload) consumes it again. */
const openPathByWc = new Map<number, string>()
/** File paths granted to each view — readFile/save only allow these */
const allowedByWc = new Map<number, Set<string>>()
/** Current save target per view; absent = untitled document */
const savePathByWc = new Map<number, string>()
/** Last seen on-disk identity (mtime+size) of each view's save target — the
 * BUG-1654 staleness fence baseline, captured at open and after every save */
const saveStampByWc = new Map<number, FileStamp | null>()
/** Unsaved-changes flags mirrored from the renderer; drives the save prompt before closing a tab/window */
const dirtyByWc = new Set<number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
/** Resolvers for menu-triggered saves, resolved when the renderer's save invoke completes */
const saveWaiters = new Map<number, (ok: boolean) => void>()

/** Fired after a save lands on a NEW path (untitled first save / Save As) — the shell syncs tab title, recents, projects */
let fileSavedHook: ((wc: WebContents, path: string) => void) | null = null

export function setMarkdownFileSavedHook(hook: (wc: WebContents, path: string) => void): void {
  fileSavedHook = hook
}

/** Fired after a "convert & open in Docs" export — the shell routes the new
 * .docx to a docs tab in the exporting view's window (senderWcId, BUG-1107) */
let docxExportedHook: ((path: string, senderWcId?: number) => void) | null = null
/** One marked cache session per app process. Old crash leftovers are removed after seven days. */
let conversionSessionPromise: Promise<string> | null = null

export function setMarkdownDocxExportedHook(
  hook: (path: string, senderWcId?: number) => void,
): void {
  docxExportedHook = hook
}

function markdownConversionSession(): Promise<string> {
  conversionSessionPromise ??= createMarkdownConversionSession(
    join(app.getPath('userData'), 'markdown-conversions'),
  )
  return conversionSessionPromise
}

/** Shell menu export entry: ask the renderer to serialize and run the export flow */
export function sendMarkdownExportRequest(contents: WebContents, format: ExportFormat): void {
  if (!contents.isDestroyed()) contents.send(MARKDOWN_CHANNELS.exportRequest, format)
}

/** Shell menu Print: ask the renderer to build the print HTML and open the system dialog */
export function sendMarkdownPrintRequest(contents: WebContents): void {
  if (!contents.isDestroyed()) contents.send(MARKDOWN_CHANNELS.printRequest)
}

export function markdownIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

// ── Crash recovery: dirty renderers push a copy every 30s
// (markdown:write-recovery); a normal save cleans it up; open offers Restore/Discard ──

const readTextDecoded = async (path: string) =>
  decodeTextBytes(await readFile(path), legacyCharsetForLang(getUiLang()))

const recoveryStore = new TextRecoveryStore(join(app.getPath('userData'), 'markdown-autosave'), {
  write: (target, text) => atomicWriteFile(target, Buffer.from(text, 'utf8')),
  readOriginal: readTextDecoded,
})

/** Restore/Discard prompt for a recovery copy newer than the opened file */
async function promptMarkdownRecovery(
  parent: BrowserWindow | null,
): Promise<'restore' | 'discard'> {
  const options = {
    type: 'question' as const,
    buttons: [tm('autosaveRestore'), tm('autosaveDiscard')],
    defaultId: 0,
    cancelId: 1,
    message: tm('autosaveFoundTitle'),
    detail: tm('autosaveFoundBody'),
  }
  const r =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  return r.response === 0 ? 'restore' : 'discard'
}

/** drop the recovery copy for a clean tab's file (in-flight-write race guarded by the store) */
function clearMarkdownRecoveryFor(wcId: number): void {
  if (dirtyByWc.has(wcId)) return
  const path = savePathByWc.get(wcId)
  if (path) recoveryStore.clear(path)
}

export function markdownFilePath(webContentsId: number): string | undefined {
  return savePathByWc.get(webContentsId)
}

/** The file was renamed on disk — re-grant the new path and tell the renderer */
export function markdownFileRenamed(contents: WebContents, oldPath: string, newPath: string): void {
  const wcId = contents.id
  if (savePathByWc.get(wcId) === oldPath) {
    savePathByWc.set(wcId, newPath)
    // a rename keeps the file's identity — re-stamp at the new path so the
    // staleness fence keeps comparing against the same file (BUG-1654)
    saveStampByWc.set(wcId, statFileStamp(newPath))
  }
  if (openPathByWc.get(wcId) === oldPath) openPathByWc.set(wcId, newPath)
  const allowed = allowedByWc.get(wcId)
  if (allowed?.has(oldPath)) allowed.add(newPath)
  void renameOwnedAssetDocument(oldPath, newPath).catch((error) => {
    console.warn('[markdown] asset manifest rename sync failed:', error)
  })
  if (!contents.isDestroyed()) contents.send(MARKDOWN_CHANNELS.fileRenamed, newPath)
}

/**
 * Close guard: true means proceed with closing. Clean → true; dirty →
 * Save / Don't Save / Cancel. On Save, ask the renderer to serialize + write
 * and await the result; a canceled untitled-save dialog keeps the tab open.
 */
export async function requestMarkdownClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!dirtyByWc.has(contents.id) || contents.isDestroyed()) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('btnSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) {
    const documentPath = savePathByWc.get(contents.id)
    if (documentPath) {
      const discarded = await discardPendingOwnedAssets(documentPath)
      if (discarded.errors.length > 0) {
        console.warn('[markdown] pending asset discard incomplete:', discarded.errors)
      }
      // the user explicitly declined to keep the edits — the crash-recovery
      // copy must not resurrect them on the next open
      recoveryStore.clear(documentPath)
    }
    return true
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(MARKDOWN_CHANNELS.closeSaveRequest)
  })
}

/** Menu Save / Save As: ask the renderer to serialize and save; clean views resolve true immediately on plain save */
export function requestMarkdownSave(contents: WebContents, mode: SaveMode): Promise<boolean> {
  if (contents.isDestroyed()) return Promise.resolve(false)
  if (mode === 'save' && !dirtyByWc.has(contents.id) && savePathByWc.has(contents.id)) {
    return Promise.resolve(true)
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      saveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    saveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(MARKDOWN_CHANNELS.saveRequest, mode)
  })
}

async function writeTextAtomic(path: string, text: string): Promise<void> {
  await atomicWriteFile(path, Buffer.from(text, 'utf8'))
}

type ExternalChangeChoice = 'saveAs' | 'overwrite' | 'cancel'

/**
 * BUG-1654: honest prompt when an in-place save hits a file that was changed,
 * renamed, or deleted outside this window. Save As is the default (nothing can
 * be lost); Overwrite deliberately replaces the external contents; Cancel
 * abandons the write and keeps the document dirty.
 */
async function promptExternalChange(
  win: BrowserWindow | undefined,
  fileName: string,
): Promise<ExternalChangeChoice> {
  const options = {
    type: 'warning' as const,
    message: tm('externalChangeTitle'),
    detail: tm('externalChangeDetail', { name: fileName }),
    buttons: [tm('btnSaveAs'), tm('btnOverwrite'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    win && !win.isDestroyed()
      ? await dialog.showMessageBox(win, options)
      : await dialog.showMessageBox(options)
  return response === 0 ? 'saveAs' : response === 1 ? 'overwrite' : 'cancel'
}

async function resolveSaveTarget(
  e: Electron.IpcMainInvokeEvent,
  mode: SaveMode,
  suggestedName?: string,
): Promise<string | null | 'canceled'> {
  const current = savePathByWc.get(e.sender.id)
  if (mode === 'save' && current) return current
  // AI auto-naming: silent first save of an untitled document
  if (mode === 'save' && !current && suggestedName) {
    // the cap counts code points (BUG-412): slice would split a surrogate pair
    const base = truncateByCodePoints(suggestedName.replace(/[/\\:*?"<>|]/g, '_'), 80).trim()
    if (base) {
      const dir = configuredDefaultSaveDir(app)
      let target = join(dir, `${base}.md`)
      for (let n = 1; existsSync(target); n++) target = join(dir, `${base}-${n}.md`)
      return target
    }
  }
  const win =
    BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
  const defaultPath = current
    ? join(dirname(current), basename(current))
    : join(configuredDefaultSaveDir(app), `${tm('untitledFile')}.md`)
  const picked = await showSaveDialogWithMemory(dialog, win, {
    title: tm('dlgSaveTitle'),
    defaultPath,
    filters: [{ name: tm('filterMarkdown'), extensions: ['md', 'markdown'] }],
  })
  if (picked.canceled || !picked.filePath) return 'canceled'
  return picked.filePath
}

const DISPLAY_IMAGE_EXTS = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
  '.webp',
  '.svg',
  '.bmp',
  '.avif',
])

/**
 * Serves authored image paths to the editor DOM. A plain file:// <img> URL is
 * blocked whenever the renderer page is served over http (dev server), so the
 * renderer resolves images to md-asset:// instead. Only image files inside an
 * open document's directory are served.
 */
function registerImageProtocol(): void {
  protocol.handle('md-asset', async (request) => {
    let target: string
    try {
      target = decodeURIComponent(new URL(request.url).pathname)
    } catch {
      return new Response(null, { status: 400 })
    }
    if (/^\/[a-zA-Z]:\//.test(target)) target = target.slice(1)
    target = resolve(target)
    if (!DISPLAY_IMAGE_EXTS.has(extname(target).toLowerCase()) || !existsSync(target)) {
      return new Response(null, { status: 404 })
    }
    let inDocDir = false
    for (const doc of new Set([...openPathByWc.values(), ...savePathByWc.values()])) {
      const dir = resolve(dirname(doc))
      if (target === dir || !target.startsWith(dir + sep)) continue
      if (await resolveSafeRelativeImagePath(doc, relative(dir, target))) {
        inDocDir = true
        break
      }
    }
    if (!inDocDir) return new Response(null, { status: 403 })
    return net.fetch(pathToFileURL(target).toString())
  })
}

let ipcRegistered = false

function registerMarkdownIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  registerImageProtocol()

  ipcMain.handle(MARKDOWN_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.handle(MARKDOWN_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('markdown: path not granted to this view')
    }
    // A recovery copy newer than the file (crash with unsaved edits) is
    // offered as Restore/Discard before the file's own bytes are served.
    return recoveryStore.maybeRecover(path, () =>
      promptMarkdownRecovery(BrowserWindow.fromWebContents(e.sender)),
    )
  })

  // crash-recovery copy push: dirty renderers serialize and send every ~30s
  ipcMain.handle(MARKDOWN_CHANNELS.writeRecovery, async (e, path: unknown, text: unknown) => {
    if (
      typeof path !== 'string' ||
      typeof text !== 'string' ||
      !allowedByWc.get(e.sender.id)?.has(path)
    ) {
      return
    }
    mkdirSync(join(app.getPath('userData'), 'markdown-autosave'), { recursive: true })
    await recoveryStore.writeCopy(path, text)
  })

  ipcMain.handle(
    MARKDOWN_CHANNELS.save,
    async (e, request: SaveMarkdownRequest): Promise<SaveMarkdownResult> => {
      const waiter = saveWaiters.get(e.sender.id)
      saveWaiters.delete(e.sender.id)
      const done = (result: SaveMarkdownResult): SaveMarkdownResult => {
        waiter?.(result.ok && !('canceled' in result))
        return result
      }
      if (typeof request?.text !== 'string') {
        return done({ ok: false, error: 'markdown: bad save request' })
      }
      if (
        request.imageSources !== undefined &&
        (!Array.isArray(request.imageSources) ||
          request.imageSources.some((source) => typeof source !== 'string'))
      ) {
        return done({ ok: false, error: 'markdown: bad image references' })
      }
      const mode: SaveMode = request.mode === 'saveAs' ? 'saveAs' : 'save'
      const pathAtRequest = savePathByWc.get(e.sender.id)
      const pendingAtRequest = pathAtRequest
        ? await pendingOwnedAssetsForDocument(pathAtRequest)
        : []
      try {
        const suggestedName =
          typeof request.suggestedName === 'string' ? request.suggestedName : undefined
        let target = await resolveSaveTarget(e, mode, suggestedName)
        if (target === 'canceled') return done({ ok: true, canceled: true })
        if (!target) return done({ ok: false, error: 'markdown: no save target' })
        // BUG-1654 staleness fence: an in-place save must not blindly
        // resurrect a path that was renamed/deleted externally (silent fork)
        // or clobber changes another window or program wrote since the last
        // save (last-writer-wins). The stat runs as late as possible before
        // the write; a writer racing between the stat and the atomic rename
        // still wins — the fence narrows the window, it cannot close it.
        if (pathAtRequest && resolve(pathAtRequest) === resolve(target)) {
          const verdict = checkSaveStaleness(target, saveStampByWc.get(e.sender.id))
          if (verdict !== 'fresh') {
            // automatic saves are declined silently — a modal every autosave
            // tick would hold the document hostage; the doc just stays dirty
            if (request.auto === true) return done({ ok: true, canceled: true })
            const win =
              BrowserWindow.fromWebContents(e.sender) ??
              BrowserWindow.getFocusedWindow() ??
              undefined
            const choice = await promptExternalChange(win, basename(target))
            if (choice === 'cancel') return done({ ok: true, canceled: true })
            if (choice === 'saveAs') {
              const retry = await resolveSaveTarget(e, 'saveAs', suggestedName)
              if (retry === 'canceled') return done({ ok: true, canceled: true })
              if (!retry) return done({ ok: false, error: 'markdown: no save target' })
              target = retry
            }
            // 'overwrite': the user deliberately replaces the external contents
          }
        }
        const currentPath = pathAtRequest
        const isNewPath = currentPath !== target
        const imageSources = [...(request.imageSources ?? [])]
        const knownImageSources = new Set(imageSources)
        for (const source of extractMarkdownImageSources(request.text)) {
          if (knownImageSources.has(source)) continue
          knownImageSources.add(source)
          imageSources.push(source)
        }
        const prepared =
          currentPath && resolve(dirname(currentPath)) !== resolve(dirname(target))
            ? await prepareAssetsForSaveAs(currentPath, target, request.text, imageSources)
            : null
        const textToWrite = prepared?.text ?? request.text
        const savedImageSources = prepared?.imageSources ?? imageSources
        try {
          await writeTextAtomic(target, textToWrite)
        } catch (error) {
          if (prepared) await rollbackPreparedSaveAsAssets(prepared).catch(() => {})
          throw error
        }
        savePathByWc.set(e.sender.id, target)
        // the write is ours — refresh the fence baseline from the fresh file
        saveStampByWc.set(e.sender.id, statFileStamp(target))
        // keep the reload path in sync — a stale openPathByWc would make a
        // reloaded renderer load the OLD file and then save it over the new one
        openPathByWc.set(e.sender.id, target)
        const allowed = allowedByWc.get(e.sender.id) ?? new Set<string>()
        allowed.add(target)
        allowedByWc.set(e.sender.id, allowed)
        dirtyByWc.delete(e.sender.id)
        const pendingNames = prepared
          ? prepared.created.map((record) => record.name)
          : currentPath && resolve(currentPath) === resolve(target)
            ? pendingAtRequest
            : []
        const reconciled = await reconcileOwnedAssets(target, savedImageSources, { pendingNames })
        if (reconciled.errors.length > 0) {
          console.warn('[markdown] asset reconciliation incomplete:', reconciled.errors)
        }
        if (mode === 'saveAs' && currentPath && resolve(currentPath) !== resolve(target)) {
          const sourceResolved = await resolveSourcePendingAfterSaveAs(
            currentPath,
            pendingAtRequest,
          )
          if (sourceResolved.errors.length > 0) {
            console.warn(
              '[markdown] source asset reconciliation incomplete:',
              sourceResolved.errors,
            )
          }
        }
        // the persisted file now carries these edits — its recovery copy must
        // not re-offer them; a save-as also retires the old file's copy
        recoveryStore.clear(target)
        if (currentPath && resolve(currentPath) !== resolve(target))
          recoveryStore.clear(currentPath)
        if (isNewPath) fileSavedHook?.(e.sender, target)
        return done({
          ok: true,
          path: target,
          ...(prepared?.rewrites.length ? { imageRewrites: prepared.rewrites } : {}),
        })
      } catch (err) {
        return done({ ok: false, error: err instanceof Error ? err.message : String(err) })
      }
    },
  )

  ipcMain.handle(MARKDOWN_CHANNELS.pickImage, async (e): Promise<string | null> => {
    const docPath = savePathByWc.get(e.sender.id)
    if (!docPath) return null
    const win =
      BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
    const picked = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgPickImage'),
      // only formats readImage/DOCX export can round-trip (docx-engine NewImage mimes)
      filters: [{ name: tm('filterImages'), extensions: ['png', 'jpg', 'jpeg', 'gif'] }],
      properties: ['openFile'],
    })
    const source = picked.filePaths[0]
    if (picked.canceled || !source) return null
    return copyImageIntoOwnedAssets(docPath, source)
  })

  ipcMain.handle(
    MARKDOWN_CHANNELS.saveImage,
    async (e, data: { base64?: unknown; ext?: unknown }): Promise<string | null> => {
      const docPath = savePathByWc.get(e.sender.id)
      const ext = String(data?.ext ?? '').toLowerCase()
      if (!docPath || typeof data?.base64 !== 'string' || !data.base64) return null
      // keep in sync with readImage's MIME map — every authored asset must stay DOCX-exportable
      if (!['png', 'jpg', 'jpeg', 'gif'].includes(ext)) return null
      return writeImageIntoOwnedAssets(docPath, `image.${ext}`, Buffer.from(data.base64, 'base64'))
    },
  )

  // markdown-owned (like docs:ai-generate-image): the shared ai:* handlers are
  // shell-registered, but image generation is gated per app
  ipcMain.handle(
    MARKDOWN_CHANNELS.aiGenerateImage,
    (_e, op: { prompt?: unknown; aspectRatio?: unknown }) =>
      generateImageTool(join(app.getPath('userData'), 'ai-settings.json'), {
        prompt: String(op?.prompt ?? ''),
        aspectRatio: op?.aspectRatio ? String(op.aspectRatio) : undefined,
      }),
  )

  const MIME_BY_EXT: Record<string, ImageData['mime']> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
  }

  ipcMain.handle(
    MARKDOWN_CHANNELS.readImage,
    async (e, src: unknown): Promise<ImageData | null> => {
      const docPath = savePathByWc.get(e.sender.id)
      if (!docPath || typeof src !== 'string' || /^[a-z][a-z0-9+.-]*:/i.test(src)) return null
      const target = await resolveSafeRelativeImagePath(docPath, src)
      if (!target) return null
      const mime = MIME_BY_EXT[extname(target).toLowerCase()]
      if (!mime || !existsSync(target)) return null
      try {
        return { base64: (await readFile(target)).toString('base64'), mime }
      } catch {
        return null
      }
    },
  )

  ipcMain.handle(
    MARKDOWN_CHANNELS.exportDocx,
    async (e, request: ExportDocxRequest): Promise<ExportResult> => {
      if (typeof request?.base64 !== 'string' || !request.base64) {
        return { ok: false, error: 'markdown: bad export request' }
      }
      // the cap counts code points (BUG-412): slice would split a surrogate pair
      const safeName =
        truncateByCodePoints(
          String(request.suggestedName || tm('untitledFile')).replace(/[/\\:*?"<>|]/g, '_'),
          80,
        ).trim() || tm('untitledFile')
      try {
        const bytes = Buffer.from(request.base64, 'base64')
        if (request.mode === 'openInDocs') {
          // Each conversion gets an app-owned cache file. A marked session is
          // retained for the life of open Docs tabs; crash leftovers expire
          // after the explicit TTL enforced when the next session starts.
          const target = await writeMarkdownConversion(
            await markdownConversionSession(),
            safeName,
            bytes,
          )
          docxExportedHook?.(target, e.sender.id)
          return { ok: true, path: target }
        }
        const win =
          BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
        const picked = await showSaveDialogWithMemory(
          dialog,
          win,
          {
            defaultPath: `${safeName}.docx`,
            filters: [{ name: 'Word', extensions: ['docx'] }],
          },
          configuredDefaultSaveDir(app),
        )
        if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
        await atomicWriteFile(picked.filePath, bytes)
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    MARKDOWN_CHANNELS.exportPdf,
    async (e, request: ExportPdfRequest): Promise<ExportResult> => {
      if (typeof request?.html !== 'string' || !request.html) {
        return { ok: false, error: 'markdown: bad export request' }
      }
      // the cap counts code points (BUG-412): slice would split a surrogate pair
      const safeName =
        truncateByCodePoints(
          String(request.suggestedName || tm('untitledFile')).replace(/[/\\:*?"<>|]/g, '_'),
          80,
        ).trim() || tm('untitledFile')
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showSaveDialogWithMemory(
        dialog,
        win,
        {
          defaultPath: `${safeName}.pdf`,
          filters: [{ name: 'PDF', extensions: ['pdf'] }],
        },
        configuredDefaultSaveDir(app),
      )
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      // sheets-style: render the print HTML in a hidden scripting-disabled window
      const workDir = await mkdtemp(join(tmpdir(), 'airy-md-pdf-'))
      const printWin = new BrowserWindow({
        show: false,
        webPreferences: { sandbox: true, javascript: false },
      })
      try {
        const htmlPath = join(workDir, 'print.html')
        await writeFile(htmlPath, request.html, 'utf8')
        await printWin.loadFile(htmlPath)
        const pdf = await printWin.webContents.printToPDF({
          pageSize: 'A4',
          printBackground: true,
          margins: { top: 0.6, bottom: 0.6, left: 0.6, right: 0.6 },
        })
        await atomicWriteFile(picked.filePath, pdf)
        openExportedPdf(picked.filePath, e.sender.id)
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      } finally {
        printWin.destroy()
        await rm(workDir, { recursive: true, force: true })
      }
    },
  )

  ipcMain.on(MARKDOWN_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(MARKDOWN_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // safety net for menu saves the renderer declined without invoking save()
  // (busy / still loading) — the save handler itself resolves the normal path
  ipcMain.on(MARKDOWN_CHANNELS.saveRequestAck, (e, ok: unknown) => {
    const waiter = saveWaiters.get(e.sender.id)
    saveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // Language channel shared with other modules; removeHandler tolerates duplicate registration
  ipcMain.removeHandler(MARKDOWN_CHANNELS.getLanguage)
  ipcMain.handle(MARKDOWN_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath)) {
    openPathByWc.set(wcId, openPath)
    savePathByWc.set(wcId, openPath)
    // fence baseline at open (BUG-1654): the renderer is about to read this file
    saveStampByWc.set(wcId, statFileStamp(openPath))
    allowedByWc.set(wcId, new Set([openPath]))
  }
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    clearMarkdownRecoveryFor(wcId)
    openPathByWc.delete(wcId)
    allowedByWc.delete(wcId)
    savePathByWc.delete(wcId)
    saveStampByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
    saveWaiters.get(wcId)?.(false)
    saveWaiters.delete(wcId)
  })
}

export function createMarkdownView(openPath?: string | null): WebContentsView {
  registerMarkdownIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  grantAndTrack(view.webContents, openPath)
  if (runtime.rendererUrl)
    voidLoad(view.webContents.loadURL(runtime.rendererUrl), 'markdown tab renderer')
  else if (runtime.rendererFile)
    voidLoad(view.webContents.loadFile(runtime.rendererFile), 'markdown tab renderer')
  return view
}

/** Standalone window mode: `npm run dev -w @airy-office/markdown`, md path passed via argv */
export function startMarkdownStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configureMarkdownRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerMarkdownIpc()
    const win = new BrowserWindow({
      width: 1200,
      height: 850,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    const argPath = process.argv.slice(1).find((a) => /\.(md|markdown)$/i.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) voidLoad(win.loadURL(runtime.rendererUrl), 'markdown window')
    else if (runtime.rendererFile) voidLoad(win.loadFile(runtime.rendererFile), 'markdown window')
  })
  app.on('window-all-closed', () => app.quit())
}
