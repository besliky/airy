import type { zh } from './zh'

export const th = {
  today: 'วันนี้',
  yesterday: 'เมื่อวาน',
  daysAgo: '{n} วันที่แล้ว',
  closeTab: 'ปิดแท็บ',
  errFileNotFound: 'ไฟล์ไม่มีอยู่หรือถูกย้ายไปแล้ว',
  errPermissionDenied: 'ไม่มีสิทธิ์เข้าถึงไฟล์นี้',
  errFileLocked: 'ไฟล์ถูกเปิดในโปรแกรมอื่นอยู่ ปิดแล้วลองอีกครั้ง',
  errTooManyFiles: 'เปิดไฟล์มากเกินไป โปรดลองอีกครั้งในอีกครู่',
  tabList: 'แท็บทั้งหมด',
  newTab: 'แท็บใหม่',
} satisfies Record<keyof typeof zh, string>
