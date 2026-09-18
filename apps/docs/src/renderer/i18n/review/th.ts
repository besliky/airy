import type { zh } from './zh'

export const th = {
  reviewCompareMerge: 'เปรียบเทียบ (รวมเป็นการแก้ไข)',
  reviewCompareMergeDesc: 'แสดงความต่างเป็นการแก้ไขที่ติดตามซึ่งยอมรับหรือปฏิเสธได้',
  reviewComparePanel: 'แสดงเฉพาะแผงความแตกต่าง',
  reviewComparePanelDesc: 'แสดงรายการความต่างของย่อหน้าในบานหน้าต่างด้านข้างโดยไม่แก้ไขเอกสาร',
  reviewCompareMerged:
    'เปรียบเทียบกับ {name} แล้ว: แทรก {added} จุด ลบ {removed} จุด แก้ไข {changed} จุด ถูกรวมเป็นการติดตามการแก้ไข',
  reviewCompareIdentical: 'ไม่มีความแตกต่างกับ {name}: เอกสารทั้งสองเหมือนกัน',
  reviewCompareMergedApprox:
    'เปรียบเทียบกับ {name} แล้ว: แทรก {added} จุด ลบ {removed} จุด แก้ไข {changed} จุด ถูกรวมเป็นการติดตามการแก้ไข (เอกสารใหญ่เกินกว่าจะจับคู่ย่อหน้าได้อย่างแม่นยำ)',
  reviewCompareDegraded:
    'เอกสารใหญ่เกินกว่าจะจับคู่ย่อหน้าได้อย่างแม่นยำ: ความแตกต่างถูกจับคู่ตามตำแหน่ง',
  reviewComparePendingRevisions:
    'เอกสารมีการติดตามการแก้ไขที่ยังค้างอยู่ โปรดยอมรับหรือปฏิเสธก่อนเปรียบเทียบอีกครั้ง',
  reviewCompareReadonly:
    'การเปรียบเทียบ (รวมเป็นการแก้ไข) ต้องใช้เอกสารที่แก้ไขได้ เอกสารนี้เป็นแบบอ่านอย่างเดียว',
} satisfies Record<keyof typeof zh, string>
