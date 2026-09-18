import type { zh } from './zh'

export const th = {
  reviewCompareMerge: 'เปรียบเทียบ (รวมเป็นการแก้ไข)',
  reviewComparePanel: 'แสดงเฉพาะแผงความแตกต่าง',
  reviewCompareMerged:
    'เปรียบเทียบกับ {name} แล้ว: แทรก {added} จุด ลบ {removed} จุด แก้ไข {changed} จุด ถูกรวมเป็นการติดตามการแก้ไข',
  reviewCompareIdentical: 'ไม่มีความแตกต่างกับ {name}: เอกสารทั้งสองเหมือนกัน',
} satisfies Record<keyof typeof zh, string>
