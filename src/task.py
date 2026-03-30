# src/task.py
import logging
from pathlib import Path
from pptx.slide import Slide

from .core.ppt_tools.image_handler import replace_image_by_name
from .core.ppt_tools.table_handler import fill_rain_table
from .core.ppt_tools.text_handler import (
    replace_text_by_name,
    get_buddhist_year,
    get_thai_month,
    get_next_months,
    format_month_range,
)

logger = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# Internal helpers
# ──────────────────────────────────────────────
_SLOTS = ["t1", "t2", "t3", "t4", "t5", "t6"]


def _update_month_labels(slide: Slide, months: list) -> bool:
    """Update txt_month_t1…t6 from a pre-computed months list."""
    ok = True
    for slot, m in zip(_SLOTS, months):
        if not replace_text_by_name(slide, f"txt_month_{slot}", m["thai_name"]):
            ok = False
    return ok


def _update_6month_texts(slide: Slide, init_year: int, init_month: int) -> bool:
    """Update txt_header (month range only) and txt_month_t1…t6."""
    months = get_next_months(init_year, init_month, 6)
    ok = replace_text_by_name(slide, "txt_header", format_month_range(months))
    if not _update_month_labels(slide, months):
        ok = False
    return ok


def _replace_6_images(slide: Slide, prefix: str, paths: list) -> bool:
    """Replace pic_{prefix}_t1…t6 from the provided path list (t1 first)."""
    ok = True
    for slot, path in zip(_SLOTS, paths):
        if not replace_image_by_name(slide, f"pic_{prefix}_{slot}", path):
            ok = False
    return ok


def _log_result(tag: str, ok: bool) -> None:
    if ok:
        logger.info(f"[{tag}] updated successfully.")
    else:
        logger.warning(f"[{tag}] finished with one or more failures — check logs above.")


# ──────────────────────────────────────────────
# tag_fcst_yearly — yearly forecast
# ──────────────────────────────────────────────
def update_yearly_forecast(
    slide: Slide,
    target_year: int,
    path_avg30y: Path | str,
    path_fcst: Path | str,
    path_anom: Path | str,
) -> bool:
    """tag_fcst_yearly — คาดการณ์ฝนรายปี"""
    thai_year = get_buddhist_year(target_year)
    logger.info(f"Updating yearly forecast: {thai_year}")
    ok = True

    for shape_name, text in {
        "txt_header":          f"คาดการณ์ฝนปี {thai_year}",
        "txt_label_fcst_year": f"คาดการณ์ ปี {thai_year}",
    }.items():
        if not replace_text_by_name(slide, shape_name, text):
            ok = False

    for shape_name, path in {
        "pic_avg_yearly":       path_avg30y,
        "pic_fcst_yearly":      path_fcst,
        "pic_anom_fcst_yearly": path_anom,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False

    _log_result("tag_fcst_yearly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_hii_monthly — HII monthly forecast
# ──────────────────────────────────────────────
def update_hii_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,  # [path_t1, …, path_t6]
) -> bool:
    """tag_hii_monthly — คาดการณ์ฝนรายเดือน สสน."""
    logger.info("Updating HII monthly forecast")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "hii", paths):
        ok = False
    _log_result("tag_hii_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_hii_anom_monthly — HII anomaly monthly
# ──────────────────────────────────────────────
def update_hii_anom_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
    tbl_data: dict | None = None,
) -> bool:
    """tag_hii_anom_monthly — ความผิดปกติฝนรายเดือน สสน."""
    logger.info("Updating HII anomaly monthly")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "hii_anom", paths):
        ok = False
    if tbl_data:
        name_to_data = {r["name"]: r["values"] for r in tbl_data["rows"]}
        if not fill_rain_table(slide, "tbl_region_hii_anom", name_to_data):
            ok = False
    _log_result("tag_hii_anom_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_om_monthly — One Map mean monthly forecast
# ──────────────────────────────────────────────
def update_om_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
) -> bool:
    """tag_om_monthly — คาดการณ์ฝนรายเดือน One Map Mean"""
    logger.info("Updating One Map mean monthly forecast")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "om", paths):
        ok = False
    _log_result("tag_om_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_om_anom_monthly — One Map mean anomaly monthly
# ──────────────────────────────────────────────
def update_om_anom_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
    tbl_data: dict | None = None,
) -> bool:
    """tag_om_anom_monthly — ความผิดปกติฝนรายเดือน One Map Mean"""
    logger.info("Updating One Map mean anomaly monthly")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "om_anom", paths):
        ok = False
    if tbl_data:
        name_to_data = {r["name"]: r["values"] for r in tbl_data["rows"]}
        if not fill_rain_table(slide, "tbl_region_om_anom", name_to_data):
            ok = False
    _log_result("tag_om_anom_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_om_upper_monthly — One Map upper monthly forecast
# ──────────────────────────────────────────────
def update_om_upper_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
) -> bool:
    """tag_om_upper_monthly — คาดการณ์ฝนรายเดือน One Map Upper"""
    logger.info("Updating One Map upper monthly forecast")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "om_upper", paths):
        ok = False
    _log_result("tag_om_upper_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_om_upper_anom_monthly — One Map upper anomaly monthly
# ──────────────────────────────────────────────
def update_om_upper_anom_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
    tbl_data: dict | None = None,
) -> bool:
    """tag_om_upper_anom_monthly — ความผิดปกติฝนรายเดือน One Map Upper"""
    logger.info("Updating One Map upper anomaly monthly")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "om_upper_anom", paths):
        ok = False
    if tbl_data:
        name_to_data = {r["name"]: r["values"] for r in tbl_data["rows"]}
        if not fill_rain_table(slide, "tbl_region_om_upper_anom", name_to_data):
            ok = False
    _log_result("tag_om_upper_anom_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_om_lower_monthly — One Map lower monthly forecast
# ──────────────────────────────────────────────
def update_om_lower_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
) -> bool:
    """tag_om_lower_monthly — คาดการณ์ฝนรายเดือน One Map Lower"""
    logger.info("Updating One Map lower monthly forecast")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "om_lower", paths):
        ok = False
    _log_result("tag_om_lower_monthly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_om_lower_anom_monthly — One Map lower anomaly monthly
# ──────────────────────────────────────────────
def update_om_lower_anom_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths: list,
    tbl_data: dict | None = None,
) -> bool:
    """tag_om_lower_anom_monthly — ความผิดปกติฝนรายเดือน One Map Lower"""
    logger.info("Updating One Map lower anomaly monthly")
    ok = _update_6month_texts(slide, init_year, init_month)
    if not _replace_6_images(slide, "om_lower_anom", paths):
        ok = False
    if tbl_data:
        name_to_data = {r["name"]: r["values"] for r in tbl_data["rows"]}
        if not fill_rain_table(slide, "tbl_region_om_lower_anom", name_to_data):
            ok = False
    _log_result("tag_om_lower_anom_monthly", ok)
    return ok


# ══════════════════════════════════════════════════════════════
# Group 2.10 — Observed vs Forecast / Avg30y
# Each function receives a single slide + pre-resolved paths.
# txt_header is built from obs_year / obs_month.
# rect_val_* shapes (rainfall values from extract_rain_to_excel)
# are intentionally skipped until that integration is complete.
# ══════════════════════════════════════════════════════════════

# ──────────────────────────────────────────────
# tag_obs_vs_hii_yearly — January edition only
# Shapes: pic_fcst, pic_obs, pic_diff, txt_header
# ──────────────────────────────────────────────
def update_obs_vs_hii_yearly(
    slide: Slide,
    obs_year: int,       # the year being reviewed (e.g. 2025 in a Jan-2026 report)
    path_fcst: Path | str,   # HII yearly forecast image
    path_obs: Path | str,    # yearly observed image
    path_diff: Path | str,   # yearly diff (HII fcst – obs)
) -> bool:
    """tag_obs_vs_hii_yearly — เปรียบเทียบฝนตรวจวัดรายปี กับคาดการณ์ สสน. (ฉบับ ม.ค.)"""
    thai_year = get_buddhist_year(obs_year)
    logger.info(f"Updating obs_vs_hii_yearly: {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบฝนตรวจวัดปี {thai_year} กับคาดการณ์ฝน สสน.",
    )
    for shape_name, path in {
        "pic_fcst": path_fcst,
        "pic_obs":  path_obs,
        "pic_diff": path_diff,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result("tag_obs_vs_hii_yearly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_obs_vs_hii — monthly obs vs HII forecast
# Shapes: pic_fcst, pic_obs, pic_diff, txt_header
# Note: per slide_notes 2.10.1, pic_fcst holds the monthly
#       avg30y reference image (not the HII forecast itself).
# ──────────────────────────────────────────────
def update_obs_vs_hii(
    slide: Slide,
    obs_year: int,
    obs_month: int,
    path_fcst: Path | str,   # avg30y monthly (reference context image)
    path_obs: Path | str,    # observed monthly
    path_diff: Path | str,   # diff: HII fcst – observed
) -> bool:
    """tag_obs_vs_hii — เปรียบเทียบฝนตรวจวัดรายเดือน กับคาดการณ์ สสน."""
    thai_month = get_thai_month(obs_month)
    thai_year  = get_buddhist_year(obs_year)
    logger.info(f"Updating obs_vs_hii: {thai_month} {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบฝนตรวจวัด เดือน{thai_month} {thai_year} กับคาดการณ์ฝน สสน.",
    )
    for shape_name, path in {
        "pic_fcst": path_fcst,
        "pic_obs":  path_obs,
        "pic_diff": path_diff,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result("tag_obs_vs_hii", ok)
    return ok


# ──────────────────────────────────────────────
# tag_obs_vs_tmd — monthly obs vs TMD forecast
# Shapes: pic_fcst, pic_obs, pic_diff, txt_header
# ──────────────────────────────────────────────
def update_obs_vs_tmd(
    slide: Slide,
    obs_year: int,
    obs_month: int,
    path_fcst: Path | str,   # TMD forecast monthly
    path_obs: Path | str,    # observed monthly
    path_diff: Path | str,   # diff: TMD fcst – observed
) -> bool:
    """tag_obs_vs_tmd — เปรียบเทียบฝนตรวจวัดรายเดือน กับคาดการณ์ กรมอุตุฯ"""
    thai_month = get_thai_month(obs_month)
    thai_year  = get_buddhist_year(obs_year)
    logger.info(f"Updating obs_vs_tmd: {thai_month} {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบฝนตรวจวัด เดือน{thai_month} {thai_year} กับคาดการณ์ฝน กรมอุตุฯ",
    )
    for shape_name, path in {
        "pic_fcst": path_fcst,
        "pic_obs":  path_obs,
        "pic_diff": path_diff,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result("tag_obs_vs_tmd", ok)
    return ok


# ──────────────────────────────────────────────
# tag_obs_vs_om — monthly obs vs One Map mean forecast
# Shapes: pic_fcst, pic_obs, pic_diff, txt_header
# ──────────────────────────────────────────────
def update_obs_vs_om(
    slide: Slide,
    obs_year: int,
    obs_month: int,
    path_fcst: Path | str,   # One Map Mean forecast monthly
    path_obs: Path | str,    # observed monthly
    path_diff: Path | str,   # diff: OM fcst – observed
) -> bool:
    """tag_obs_vs_om — เปรียบเทียบฝนตรวจวัดรายเดือน กับคาดการณ์ One Map"""
    thai_month = get_thai_month(obs_month)
    thai_year  = get_buddhist_year(obs_year)
    logger.info(f"Updating obs_vs_om: {thai_month} {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบฝนตรวจวัด เดือน{thai_month} {thai_year} กับคาดการณ์ฝน One Map",
    )
    for shape_name, path in {
        "pic_fcst": path_fcst,
        "pic_obs":  path_obs,
        "pic_diff": path_diff,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result("tag_obs_vs_om", ok)
    return ok


# ──────────────────────────────────────────────
# tag_obs_vs_avg_yearly — January edition only
# Shapes: pic_avg, pic_obs, pic_diff, txt_header
# ──────────────────────────────────────────────
def update_obs_vs_avg_yearly(
    slide: Slide,
    obs_year: int,
    path_avg: Path | str,    # avg30y yearly image
    path_obs: Path | str,    # yearly observed image
    path_diff: Path | str,   # yearly diff: observed – avg30y
) -> bool:
    """tag_obs_vs_avg_yearly — เปรียบเทียบฝนตรวจวัดรายปี กับค่าเฉลี่ย 30 ปี (ฉบับ ม.ค.)"""
    thai_year = get_buddhist_year(obs_year)
    logger.info(f"Updating obs_vs_avg_yearly: {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบฝนตรวจวัดปี {thai_year} กับค่าเฉลี่ย 30 ปี",
    )
    for shape_name, path in {
        "pic_avg":  path_avg,
        "pic_obs":  path_obs,
        "pic_diff": path_diff,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result("tag_obs_vs_avg_yearly", ok)
    return ok


# ──────────────────────────────────────────────
# tag_obs_vs_avg — monthly obs vs avg30y
# Shapes: pic_avg, pic_obs, pic_diff, txt_header
# ──────────────────────────────────────────────
def update_obs_vs_avg(
    slide: Slide,
    obs_year: int,
    obs_month: int,
    path_avg: Path | str,    # avg30y monthly image
    path_obs: Path | str,    # observed monthly
    path_diff: Path | str,   # diff: observed – avg30y
) -> bool:
    """tag_obs_vs_avg — เปรียบเทียบฝนตรวจวัดรายเดือน กับค่าเฉลี่ย 30 ปี"""
    thai_month = get_thai_month(obs_month)
    thai_year  = get_buddhist_year(obs_year)
    logger.info(f"Updating obs_vs_avg: {thai_month} {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบฝนตรวจวัด เดือน{thai_month} {thai_year} กับค่าเฉลี่ย 30 ปี",
    )
    for shape_name, path in {
        "pic_avg":  path_avg,
        "pic_obs":  path_obs,
        "pic_diff": path_diff,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result("tag_obs_vs_avg", ok)
    return ok


# ══════════════════════════════════════════════════════════════
# Group 2.11 — Forecast vs avg30y  (tag_fcst_vs_avg_t1 … t6)
# One fixed slide per forecast month — no cloning needed.
# ══════════════════════════════════════════════════════════════

def update_fcst_vs_avg(
    slide: Slide,
    target_year: int,
    target_month: int,
    path_avg: Path | str,        # avg30y monthly (region)
    path_tmd: Path | str,        # TMD forecast
    path_tmd_anom: Path | str,   # TMD forecast anomaly (diff from avg30y)
    path_hii: Path | str,        # HII forecast
    path_hii_anom: Path | str,   # HII forecast anomaly (diff from avg30y)
) -> bool:
    """tag_fcst_vs_avg_t{n} — เปรียบเทียบคาดการณ์ฝนรายเดือน กับค่าเฉลี่ย 30 ปี"""
    thai_month = get_thai_month(target_month)
    thai_year  = get_buddhist_year(target_year)
    logger.info(f"Updating fcst_vs_avg: {thai_month} {thai_year}")
    ok = replace_text_by_name(
        slide, "txt_header",
        f"เปรียบเทียบคาดการณ์ฝนเดือน{thai_month} {thai_year} กับค่าเฉลี่ย 30 ปี (2534–2563)",
    )
    for shape_name, path in {
        "pic_avg":      path_avg,
        "pic_tmd":      path_tmd,
        "pic_tmd_anom": path_tmd_anom,
        "pic_hii":      path_hii,
        "pic_hii_anom": path_hii_anom,
    }.items():
        if not replace_image_by_name(slide, shape_name, path):
            ok = False
    _log_result(f"tag_fcst_vs_avg ({thai_month} {thai_year})", ok)
    return ok


# ══════════════════════════════════════════════════════════════
# Group 2.12 — Basin-level forecast pages
#
# Image pages: 12 images each (6 forecast + 6 anomaly, basin area)
#              txt_header holds the FULL title including model name.
# Table pages: txt_header only — tables skipped until
#              extract_rain_to_excel integration is complete.
# ══════════════════════════════════════════════════════════════

# ── Internal helper ───────────────────────────────────────────
def _update_basin_image_slide(
    slide: Slide,
    header: str,
    fcst_prefix: str,
    anom_prefix: str,
    init_year: int,
    init_month: int,
    paths_fcst: list,
    paths_anom: list,
) -> bool:
    months = get_next_months(init_year, init_month, 6)
    ok = replace_text_by_name(slide, "txt_header", header)
    if not _update_month_labels(slide, months):
        ok = False
    if not _replace_6_images(slide, fcst_prefix, paths_fcst):
        ok = False
    if not _replace_6_images(slide, anom_prefix, paths_anom):
        ok = False
    return ok


def _update_basin_tbl_slide(
    slide: Slide,
    header: str,
    left_shape: str,
    right_shape: str,
    tbl_data: dict | None,
) -> bool:
    ok = replace_text_by_name(slide, "txt_header", header)
    if tbl_data:
        name_to_data = {r["name"]: r["values"] for r in tbl_data["rows"]}
        if not fill_rain_table(slide, left_shape,  name_to_data):
            ok = False
        if not fill_rain_table(slide, right_shape, name_to_data):
            ok = False
    return ok


# ── tag_hii_basin_monthly ─────────────────────────────────────
def update_hii_basin_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths_fcst: list,   # basin HII forecast t1-t6
    paths_anom: list,   # basin HII anomaly t1-t6
) -> bool:
    """tag_hii_basin_monthly — คาดการณ์ฝน สสน. รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"คาดการณ์ฝน สสน. และผลต่างจากค่าเฉลี่ย 30 ปี (2534-2563) "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating HII basin monthly")
    ok = _update_basin_image_slide(
        slide, header, "hii_fcst", "hii_anom", init_year, init_month, paths_fcst, paths_anom,
    )
    _log_result("tag_hii_basin_monthly", ok)
    return ok


# ── tag_hii_basin_tbl ─────────────────────────────────────────
def update_hii_basin_tbl(
    slide: Slide,
    init_year: int,
    init_month: int,
    tbl_data: dict | None = None,
) -> bool:
    """tag_hii_basin_tbl — ตารางฝนคาดการณ์ สสน. รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"ตารางแสดงผลต่างจากค่าปกติของปริมาณฝนคาดการณ์ สสน. "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating HII basin table")
    ok = _update_basin_tbl_slide(slide, header, "tbl_basin_hii_left", "tbl_basin_hii_right", tbl_data)
    _log_result("tag_hii_basin_tbl", ok)
    return ok


# ── tag_om_basin_monthly ──────────────────────────────────────
def update_om_basin_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths_fcst: list,
    paths_anom: list,
) -> bool:
    """tag_om_basin_monthly — คาดการณ์ฝน One Map รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"คาดการณ์ฝน One Map และผลต่างจากค่าเฉลี่ย 30 ปี (2534-2563) "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating OM basin monthly")
    ok = _update_basin_image_slide(
        slide, header, "om_fcst", "om_anom", init_year, init_month, paths_fcst, paths_anom,
    )
    _log_result("tag_om_basin_monthly", ok)
    return ok


# ── tag_om_basin_tbl ──────────────────────────────────────────
def update_om_basin_tbl(
    slide: Slide,
    init_year: int,
    init_month: int,
    tbl_data: dict | None = None,
) -> bool:
    """tag_om_basin_tbl — ตารางฝนคาดการณ์ One Map รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"ตารางแสดงผลต่างจากค่าปกติของปริมาณฝนคาดการณ์ One Map "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating OM basin table")
    ok = _update_basin_tbl_slide(slide, header, "tbl_basin_om_left", "tbl_basin_om_right", tbl_data)
    _log_result("tag_om_basin_tbl", ok)
    return ok


# ── tag_om_upper_basin_monthly ────────────────────────────────
def update_om_upper_basin_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths_fcst: list,
    paths_anom: list,
) -> bool:
    """tag_om_upper_basin_monthly — คาดการณ์ฝนสูงสุด One Map รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"คาดการณ์ฝนสูงสุด One Map และผลต่างจากค่าเฉลี่ย 30 ปี (2534-2563) "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating OM upper basin monthly")
    ok = _update_basin_image_slide(
        slide, header, "om_upper_fcst", "om_upper_anom", init_year, init_month, paths_fcst, paths_anom,
    )
    _log_result("tag_om_upper_basin_monthly", ok)
    return ok


# ── tag_om_upper_basin_tbl ────────────────────────────────────
def update_om_upper_basin_tbl(
    slide: Slide,
    init_year: int,
    init_month: int,
    tbl_data: dict | None = None,
) -> bool:
    """tag_om_upper_basin_tbl — ตารางฝนคาดการณ์สูงสุด One Map รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"ตารางแสดงผลต่างจากค่าปกติของปริมาณฝนคาดการณ์สูงสุด One Map "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating OM upper basin table")
    ok = _update_basin_tbl_slide(slide, header, "tbl_basin_om_upper_left", "tbl_basin_om_upper_right", tbl_data)
    _log_result("tag_om_upper_basin_tbl", ok)
    return ok


# ── tag_om_lower_basin_monthly ────────────────────────────────
def update_om_lower_basin_monthly(
    slide: Slide,
    init_year: int,
    init_month: int,
    paths_fcst: list,
    paths_anom: list,
) -> bool:
    """tag_om_lower_basin_monthly — คาดการณ์ฝนต่ำสุด One Map รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"คาดการณ์ฝนต่ำสุด One Map และผลต่างจากค่าเฉลี่ย 30 ปี (2534-2563) "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating OM lower basin monthly")
    ok = _update_basin_image_slide(
        slide, header, "om_lower_fcst", "om_lower_anom", init_year, init_month, paths_fcst, paths_anom,
    )
    _log_result("tag_om_lower_basin_monthly", ok)
    return ok


# ── tag_om_lower_basin_tbl ────────────────────────────────────
def update_om_lower_basin_tbl(
    slide: Slide,
    init_year: int,
    init_month: int,
    tbl_data: dict | None = None,
) -> bool:
    """tag_om_lower_basin_tbl — ตารางฝนคาดการณ์ต่ำสุด One Map รายลุ่มน้ำ"""
    months = get_next_months(init_year, init_month, 6)
    header = (
        f"ตารางแสดงผลต่างจากค่าปกติของปริมาณฝนคาดการณ์ต่ำสุด One Map "
        f"รายลุ่มน้ำ เดือน{format_month_range(months)}"
    )
    logger.info("Updating OM lower basin table")
    ok = _update_basin_tbl_slide(slide, header, "tbl_basin_om_lower_left", "tbl_basin_om_lower_right", tbl_data)
    _log_result("tag_om_lower_basin_tbl", ok)
    return ok
