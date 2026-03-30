# src/core/ppt_tools/image_handler.py
import logging
from pathlib import Path
from pptx.slide import Slide
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)

def _find_shape_recursive(shapes, shape_name: str):
    """
    ฟังก์ชันภายใน: ค้นหา Shape ตามชื่อแบบเจาะลึกเข้าไปใน Group
    """
    for shape in shapes:
        if shape.name == shape_name:
            return shape
        # ถ้าเจอ Group, ให้เจาะเข้าไปหารอบใหม่
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            found_shape = _find_shape_recursive(shape.shapes, shape_name)
            if found_shape:
                return found_shape
    return None

def replace_image_by_name(slide: Slide, shape_name: str, new_image_path: Path | str) -> bool:
    """
    ค้นหาและแทนที่รูปภาพตามชื่อ โดยคงตำแหน่ง ขนาด และสัดส่วนเดิมไว้
    * รองรับรูปภาพเดี่ยว และรูปภาพที่ถูกจัดกลุ่ม (Group Shape) *
    * หมายเหตุ: รูปใหม่จะถูกนำมาวางไว้ชั้นบนสุด (Layering) *
    """
    # 1. ตรวจสอบไฟล์รูปภาพใหม่
    if not new_image_path:
        logger.warning(f"ไม่ได้ระบุ Path สำหรับรูปภาพใหม่ ของ Shape: '{shape_name}'")
        return False
        
    path_obj = Path(new_image_path)
    if not path_obj.exists():
        logger.error(f"ไม่พบไฟล์รูปภาพใหม่ตาม Path: {new_image_path} (สำหรับ Shape: '{shape_name}')")
        return False

    # 2. ค้นหา Shape เป้าหมาย (แบบเจาะลึก Group)
    target_shape = _find_shape_recursive(slide.shapes, shape_name)
    
    # 3. ตรวจสอบ Error (ตามลอจิกไฟล์ตัวอย่างที่คุณส่งมา)
    if not target_shape:
        logger.warning(f"ไม่พบ Object ชื่อ '{shape_name}' บนสไลด์หน้า {slide.slide_id}")
        return False

    if target_shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        logger.warning(f"Object ชื่อ '{shape_name}' ถูกพบ แต่ไม่ใช่รูปภาพ (เป็นชนิด: {target_shape.shape_type})")
        return False

    logger.info(f"กำลังแทนที่รูปภาพ: '{shape_name}' ด้วยไฟล์: {path_obj.name}")

    # 4. เก็บค่าตำแหน่งและขนาดเดิม
    left = target_shape.left
    top = target_shape.top
    width = target_shape.width
    height = target_shape.height

    # 5. ลบรูปเดิมออก (Delete old shape)
    # python-pptx ลบ shape โดยตรงไม่ได้ ต้องลบที่ระดับ element xml
    old_picture_element = target_shape.element
    old_picture_element.getparent().remove(old_picture_element)

    # 6. แทรกรูปใหม่ลงไปที่ตำแหน่งและขนาดเดิม (Stretch to fit)
    # รูปใหม่จะอยู่ชั้นบนสุดโดยอัตโนมัติ
    try:
        slide.shapes.add_picture(str(path_obj), left, top, width, height)
    except Exception as e:
        logger.error(f"เกิดข้อผิดพลาดขณะแทรกรูปภาพใหม่ '{shape_name}': {e}")
        return False

    return True