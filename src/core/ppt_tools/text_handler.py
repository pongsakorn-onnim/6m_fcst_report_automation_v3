# src/core/ppt_tools/text_handler.py
import logging
from pptx.slide import Slide
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)

# ==========================================
# 1. Utilities จัดการวันที่ (จากโค้ดของคุณ)
# ==========================================
THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def get_thai_month(month_idx: int) -> str:
    """คืนค่าชื่อเดือนไทยเต็ม (1-12)"""
    if 1 <= month_idx <= 12:
        return THAI_MONTHS[month_idx - 1]
    return ""

def get_buddhist_year(year: int) -> int:
    """แปลง ค.ศ. เป็น พ.ศ."""
    return year + 543

def get_next_months(start_year: int, start_month: int, n: int):
    if n <= 0:
        return []
    results = []
    for i in range(n):
        future_month_idx = (start_month + i - 1) % 12 + 1
        year_offset = (start_month + i - 1) // 12
        future_year = start_year + year_offset
        results.append({
            "year": future_year,
            "month": future_month_idx,
            "thai_name": get_thai_month(future_month_idx),
            "buddhist_year": get_buddhist_year(future_year),
        })
    return results

def get_months_for_leads(start_year: int, start_month: int, leads: list[int]):
    if not leads:
        return []
    max_lead = max(leads)
    all_months = get_next_months(start_year, start_month, max_lead + 1)
    return [all_months[l] for l in leads]

def format_month_range(months: list[dict]) -> str:
    """Standard month-range style for ALL slides"""
    if not months:
        return ""
    start = months[0]
    end = months[-1]
    sep = " – "
    if start["buddhist_year"] == end["buddhist_year"]:
        return f"{start['thai_name']}{sep}{end['thai_name']} {start['buddhist_year']}"
    return (
        f"{start['thai_name']} {start['buddhist_year']}"
        f"{sep}"
        f"{end['thai_name']} {end['buddhist_year']}"
    )


# ==========================================
# 2. ฟังก์ชันจัดการ Text Box บน PowerPoint
# ==========================================
def _find_shape_recursive(shapes, shape_name: str):
    """ค้นหา Shape ตามชื่อแบบเจาะลึกเข้าไปใน Group"""
    for shape in shapes:
        if shape.name == shape_name:
            return shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            found_shape = _find_shape_recursive(shape.shapes, shape_name)
            if found_shape:
                return found_shape
    return None

def replace_text_by_name(slide: Slide, shape_name: str, new_text: str) -> bool:
    """
    ค้นหาและแทนที่ข้อความตามชื่อ Object โดยรักษา Format (Font, Size, Color) เดิมไว้
    *แก้ปัญหา Runs พังโดยการเก็บรูปแบบของ Run แรกไว้*
    """
    target_shape = _find_shape_recursive(slide.shapes, shape_name)

    if not target_shape:
        logger.warning(f"ไม่พบ Text Box ชื่อ '{shape_name}'")
        return False

    if not target_shape.has_text_frame:
        logger.warning(f"Object '{shape_name}' ไม่ใช่กล่องข้อความ")
        return False

    logger.info(f"อัปเดตข้อความ '{shape_name}' -> '{new_text}'")
    
    text_frame = target_shape.text_frame

    # เทคนิครักษา Format: ใส่ข้อความใหม่ลงใน Run แรกสุด แล้วเคลียร์ Run อื่นๆ ทิ้งทั้งหมด
    if text_frame.paragraphs:
        p0 = text_frame.paragraphs[0]
        if p0.runs:
            p0.runs[0].text = str(new_text)
            
            # ลบข้อความที่เหลือในบรรทัดเดียวกัน
            for run in p0.runs[1:]:
                run.text = ""
                
            # ลบข้อความในบรรทัดอื่นๆ (ถ้ามีหลายย่อหน้า)
            for paragraph in text_frame.paragraphs[1:]:
                for run in paragraph.runs:
                    run.text = ""
        else:
            # ถ้ามี Paragraph แต่ไม่มี Run เลย ให้ตั้งค่าตรงๆ
            target_shape.text = str(new_text)
    else:
        # ถ้าไม่มีอะไรเลย
        target_shape.text = str(new_text)

    return True