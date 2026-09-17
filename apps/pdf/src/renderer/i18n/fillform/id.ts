import type { en } from './en'

/** Fill Form strings (id); missing keys fall back to the en base. */
export const id = {
  ribbonTabFillForm: 'Isi formulir',
  formPreviousField: 'Kolom sebelumnya',
  formNextField: 'Kolom berikutnya',
  formFieldProgress: '{current} / {total}',
  insertText: 'Sisipkan teks',
  insertTextHint: 'Sisipkan teks yang dapat dicari ke dalam PDF',
  insertTextTitle: 'Sisipkan teks',
  editInsertedText: 'Edit teks yang disisipkan',
  deleteInsertedText: 'Hapus teks yang disisipkan',
  insertedTextDeleted: 'Teks yang disisipkan dihapus',
  textInsertSkipped: 'Teks yang disisipkan tidak dapat disimpan pada halaman: {pages}',
  textInsertNoFont:
    'Tidak ada font terpasang yang dapat menggambar teks ini dalam PDF (emoji dan simbol khusus tidak didukung)',
  formComplete: 'Selesaikan pengisian',
  formMissingRequired: '{count} kolom wajib masih kosong',
  formCompleteDone: 'Pemeriksaan pengisian formulir lolos',
  formSignField: 'Klik untuk menandatangani',
  formAddText: 'Tambahkan teks',
  formAddTextHint: 'Ketik teks, lalu klik halaman untuk menempatkannya',
  formAddTextTitle: 'Tambahkan teks ke PDF',
  formEditText: 'Edit teks',
  formAddTextPlaceholder: 'Masukkan teks yang akan ditempatkan',
  formTextSize: 'Ukuran font',
  formTextColor: 'Warna',
  formTextAlign: 'Perataan',
  formAlignLeft: 'Kiri',
  formAlignCenter: 'Tengah',
  formAlignRight: 'Kanan',
  formAddCheck: 'Centang',
  formAddCheckHint: 'Klik halaman untuk menempatkan tanda centang',
  formAddCross: 'Silang',
  formAddCrossHint: 'Klik halaman untuk menempatkan tanda X',
  formPlaceStaticHint:
    'Klik untuk menempatkan; pilih hasilnya untuk memindahkan atau mengubah ukurannya',
  formXfaWarning:
    'PDF ini berisi XFA. Hanya AcroForm yang didukung; penyimpanan mungkin tidak mempertahankan data XFA.',
} as const satisfies Partial<Record<keyof typeof en, string>>
