# src/core/output_manager.py
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

THAI_MONTHS_SHORT = [
    "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
]

@dataclass(frozen=True)
class OutputSpec:
    year: int    # ปี ค.ศ. (เช่น 2026)
    month: int   # เดือนที่ออกรายงาน (1-12)

class OutputManager:
    """
    รับผิดชอบเรื่องการตั้งชื่อไฟล์และจัดการโฟลเดอร์ปลายทาง
    รูปแบบ: YYYYMMDD_สถานการณ์น้ำและคาดการณ์ฝนเดือน{Start}-{End}{YY}.pptx
    ตัวอย่าง: 20260304_สถานการณ์น้ำและคาดการณ์ฝนเดือนมี.ค.-ส.ค.69.pptx
    """

    def __init__(self, base_output_dir: str | Path = "output"):
        self.base_dir = Path(base_output_dir)

    def build_output_path(self, spec: OutputSpec) -> Path:
        """สร้าง Path สำหรับเซฟไฟล์ พร้อมเช็คไฟล์ซ้ำ"""
        now = datetime.now()
        date_prefix = now.strftime("%Y%m%d")

        # คำนวณเดือนเริ่มต้น และเดือนสิ้นสุด (บวกไป 5 เดือน สำหรับคาดการณ์ 6 เดือน)
        start_month_idx = spec.month - 1
        end_month_idx = (start_month_idx + 5) % 12
        
        m_start = THAI_MONTHS_SHORT[start_month_idx]
        m_end = THAI_MONTHS_SHORT[end_month_idx]

        # คำนวณ พ.ศ. แบบย่อ (2026 + 543 = 2569 -> เอาแค่ "69")
        thai_year_short = str(spec.year + 543)[-2:]

        # ประกอบชื่อไฟล์
        filename = f"{date_prefix}_สถานการณ์น้ำและคาดการณ์ฝนเดือน{m_start}-{m_end}{thai_year_short}.pptx"

        # สร้างโฟลเดอร์แยกตามปีและเดือน (เช่น output/2026/03/) 
        out_dir = self.base_dir / str(spec.year) / f"{spec.month:02d}"
        out_dir.mkdir(parents=True, exist_ok=True)

        return self._get_unique_filepath(out_dir, filename)

    def _get_unique_filepath(self, directory: Path, filename: str) -> Path:
        """เช็คไฟล์ซ้ำ ถ้าซ้ำให้เติม (1), (2), (3) ... แบบ Windows"""
        name_stem = Path(filename).stem
        suffix = Path(filename).suffix
        
        counter = 1
        final_path = directory / filename
        
        while final_path.exists():
            new_name = f"{name_stem} ({counter}){suffix}"
            final_path = directory / new_name
            counter += 1
            
        return final_path