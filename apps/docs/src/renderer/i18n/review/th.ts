import type { zh } from './zh'

export const th = {
  reviewCompareMerge: 'เปรียบเทียบ (รวมเป็นการแก้ไข)',
  reviewComparePanel: 'แสดงเฉพาะแผงความแตกต่าง',
  reviewCompareMerged:
    'เปรียบเทียบกับ {name} แล้ว: แทรก {added} จุด ลบ {removed} จุด แก้ไข {changed} จุด ถูกรวมเป็นการติดตามการแก้ไข',
  reviewCompareIdentical: 'ไม่มีความแตกต่างกับ {name}: เอกสารทั้งสองเหมือนกัน',
  reviewCompareMergedApprox:
    'เปรียบเทียบกับ {name} แล้ว: แทรก {added} จุด ลบ {removed} จุด แก้ไข {changed} จุด ถูกรวมเป็นการติดตามการแก้ไข (เอกสารใหญ่เกินกว่าจะจับคู่ย่อหน้าได้อย่างแม่นยำ)',
  reviewCompareDegraded:
    'เอกสารใหญ่เกินกว่าจะจับคู่ย่อหน้าได้อย่างแม่นยำ: ความแตกต่างถูกจับคู่ตามตำแหน่ง',
  reviewComparePendingRevisions:
    'เอกสารมีการติดตามการแก้ไขที่ยังค้างอยู่ โปรดยอมรับหรือปฏิเสธก่อนเปรียบเทียบอีกครั้ง',
} satisfies Record<keyof typeof zh, string>
