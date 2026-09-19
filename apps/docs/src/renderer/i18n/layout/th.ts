import type { zh } from './zh'

export const th = {
  layoutHyphenation: 'การแบ่งคำ',
  layoutHyphNone: 'ไม่มี',
  layoutHyphManual: 'ด้วยตนเอง',
  layoutHyphAutomatic: 'อัตโนมัติ',
  layoutHyphManualDesc: 'แทรกยัติภังค์แบบมีเงื่อนไขในตำแหน่งที่ต้องการตัดคำ',
  layoutHyphManualHint: 'แบ่งคำด้วยตนเอง: กด {keys} เพื่อแทรกยัติภังค์แบบมีเงื่อนไข',
  layoutHyphSet: 'เปิดการแบ่งคำอัตโนมัติแล้ว',
  layoutHyphUnset: 'ปิดการแบ่งคำอัตโนมัติแล้ว',
  layoutSoftHyphen: 'ยัติภังค์แบบมีเงื่อนไข',
  layoutColsDialogTitle: 'คอลัมน์',
  layoutColsMore: 'คอลัมน์เพิ่มเติม…',
  layoutColOne: 'หนึ่ง',
  layoutColTwo: 'สอง',
  layoutColThree: 'สาม',
  layoutColLeft: 'ซ้าย',
  layoutColRight: 'ขวา',
  layoutColSpacing: 'ระยะห่าง (ซม.)',
  layoutColWidth: 'ความกว้าง (ซม.)',
  layoutColWidth1: 'ความกว้างคอลัมน์ 1 (ซม.)',
  layoutLineBetween: 'เส้นคั่นคอลัมน์',
} satisfies Record<keyof typeof zh, string>
