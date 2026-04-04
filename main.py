# main.py
import argparse
import logging
import sys

from src.core.config import settings
from src.core.logging_config import setup_logging
from src.core.output_manager import OutputManager, OutputSpec
from src.core.ppt_tools.text_handler import get_next_months
from src.core.rain_data_service import RainDataService, build_obs_diff_table
from src.manager import ReportManager

from rain_services.analog_year_service import AnalogYearService
from rain_services.path_builder import RainPathBuilder

from src.task import (
    update_cover_slide,
    update_summary_slide,
    update_fcst_vs_avg,
    update_yearly_forecast,
    # group 2.12
    update_hii_basin_monthly,
    update_hii_basin_tbl,
    update_om_basin_monthly,
    update_om_basin_tbl,
    update_om_upper_basin_monthly,
    update_om_upper_basin_tbl,
    update_om_lower_basin_monthly,
    update_om_lower_basin_tbl,
    update_hii_monthly,
    update_hii_anom_monthly,
    update_om_monthly,
    update_om_anom_monthly,
    update_om_upper_monthly,
    update_om_upper_anom_monthly,
    update_om_lower_monthly,
    update_om_lower_anom_monthly,
    # group 2.10
    update_obs_vs_hii_yearly,
    update_obs_vs_hii,
    update_obs_vs_tmd,
    update_obs_vs_om,
    update_obs_vs_avg_yearly,
    update_obs_vs_avg,
)

logger = logging.getLogger(__name__)


def _get_slide(report: ReportManager, tag: str):
    slide = report.get_slide_by_tag(tag)
    if slide is None:
        logger.warning(f"Slide tag '{tag}' not found — skipped.")
    return slide


def main(year: int, month: int) -> None:
    setup_logging(level="INFO", console_style="user")
    logger.info(f"Generating report: {year}-{month:02d}")

    # --- External services ---
    analog  = AnalogYearService()
    builder = RainPathBuilder(analog)

    # --- Load template ---
    report = ReportManager(settings.paths.templates_dir / "template.pptx")
    if not report.load_template():
        logger.error("Aborting: template could not be loaded.")
        sys.exit(1)

    # Pre-compute t1–t6 target months once; reused for all monthly slides
    target_months = get_next_months(year, month, 6)

    # --- Rain data service (CSV-based) ---
    rain_svc        = RainDataService(year, month)
    tbl_region_hii  = rain_svc.build_table("Region", "HII")
    tbl_region_om   = rain_svc.build_table("Region", "OM_W")
    tbl_region_om_u = rain_svc.build_table("Region", "OM_U")
    tbl_region_om_l = rain_svc.build_table("Region", "OM_L")
    tbl_basin_hii   = rain_svc.build_table("Basin",  "HII")
    tbl_basin_om    = rain_svc.build_table("Basin",  "OM_W")
    tbl_basin_om_u  = rain_svc.build_table("Basin",  "OM_U")
    tbl_basin_om_l  = rain_svc.build_table("Basin",  "OM_L")

    # ------------------------------------------------------------------
    # Page 1 — cover slide
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_cover")
    if slide:
        update_cover_slide(slide=slide, year=year, month=month)

    # ------------------------------------------------------------------
    # Page 2 — summary slide
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_summary")
    if slide:
        update_summary_slide(slide=slide, year=year, month=month)

    # ------------------------------------------------------------------
    # tag_fcst_yearly — yearly forecast
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_fcst_yearly")
    if slide:
        update_yearly_forecast(
            slide       = slide,
            target_year = year,
            path_avg30y = builder.build_avg30y_yearly(),
            path_fcst   = builder.build_hii_forecast_yearly(year, year, month),
            path_anom   = builder.build_hii_forecast_yearly(year, year, month, is_diff=True),
        )

    # ------------------------------------------------------------------
    # tag_hii_monthly — HII monthly forecast
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_hii_monthly")
    if slide:
        update_hii_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_hii_forecast_path(year, month, m["year"], m["month"])
                for m in target_months
            ],
        )

    # ------------------------------------------------------------------
    # tag_hii_anom_monthly — HII anomaly monthly
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_hii_anom_monthly")
    if slide:
        update_hii_anom_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_hii_forecast_path(year, month, m["year"], m["month"], is_diff=True)
                for m in target_months
            ],
            tbl_data   = tbl_region_hii,
        )

    # ------------------------------------------------------------------
    # tag_om_monthly — One Map mean monthly forecast
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_om_monthly")
    if slide:
        update_om_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_onemap_path(year, month, m["year"], m["month"], model_type="MFCST")
                for m in target_months
            ],
        )

    # ------------------------------------------------------------------
    # tag_om_anom_monthly — One Map mean anomaly monthly
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_om_anom_monthly")
    if slide:
        update_om_anom_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_onemap_path(year, month, m["year"], m["month"], model_type="MFCST", is_diff=True)
                for m in target_months
            ],
            tbl_data   = tbl_region_om,
        )

    # ------------------------------------------------------------------
    # tag_om_upper_monthly — One Map upper monthly forecast
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_om_upper_monthly")
    if slide:
        update_om_upper_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_onemap_path(year, month, m["year"], m["month"], model_type="UFCST")
                for m in target_months
            ],
        )

    # ------------------------------------------------------------------
    # tag_om_upper_anom_monthly — One Map upper anomaly monthly
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_om_upper_anom_monthly")
    if slide:
        update_om_upper_anom_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_onemap_path(year, month, m["year"], m["month"], model_type="UFCST", is_diff=True)
                for m in target_months
            ],
            tbl_data   = tbl_region_om_u,
        )

    # ------------------------------------------------------------------
    # tag_om_lower_monthly — One Map lower monthly forecast
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_om_lower_monthly")
    if slide:
        update_om_lower_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_onemap_path(year, month, m["year"], m["month"], model_type="LFCST")
                for m in target_months
            ],
        )

    # ------------------------------------------------------------------
    # tag_om_lower_anom_monthly — One Map lower anomaly monthly
    # ------------------------------------------------------------------
    slide = _get_slide(report, "tag_om_lower_anom_monthly")
    if slide:
        update_om_lower_anom_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths      = [
                builder.build_onemap_path(year, month, m["year"], m["month"], model_type="LFCST", is_diff=True)
                for m in target_months
            ],
            tbl_data   = tbl_region_om_l,
        )

    # ------------------------------------------------------------------
    # Group 2.10 — Observed vs Forecast / Avg30y (dynamic pages)
    # ------------------------------------------------------------------
    # January edition: compare all 12 months of the previous year.
    # Other months: compare Jan … (month-1) of the current year.
    is_january = (month == 1)
    if is_january:
        obs_year_base   = year - 1
        past_months_210 = [(obs_year_base, m) for m in range(1, 13)]
    else:
        obs_year_base   = year
        past_months_210 = [(year, m) for m in range(1, month)]

    n_obs = len(past_months_210)   # 0 if month==1 would not happen (already handled above)

    # ── Yearly slides (January edition only) ──────────────────────
    slide_hii_yearly = _get_slide(report, "tag_obs_vs_hii_yearly")
    slide_avg_yearly = _get_slide(report, "tag_obs_vs_avg_yearly")

    # ── Monthly comparison slides (cloned N times each) ────────────
    # Clone BEFORE removing yearly slides so add_slide() gets partnames
    # slide37+ and doesn't collide with still-existing slide35/slide36.
    def _build_comparison_slides(tag: str) -> list:
        tmpl = _get_slide(report, tag)
        if tmpl is None or n_obs == 0:
            return []
        slides = [tmpl]
        prev = tmpl
        for _ in range(1, n_obs):
            new_slide = report.clone_slide_after(tmpl, prev)
            slides.append(new_slide)
            prev = new_slide
        return slides

    hii_slides = _build_comparison_slides("tag_obs_vs_hii")
    tmd_slides = _build_comparison_slides("tag_obs_vs_tmd")
    om_slides  = _build_comparison_slides("tag_obs_vs_om")
    avg_slides = _build_comparison_slides("tag_obs_vs_avg")

    if not is_january:
        # Remove yearly slides — they don't belong in non-January reports
        if slide_hii_yearly:
            report.remove_slide(slide_hii_yearly)
        if slide_avg_yearly:
            report.remove_slide(slide_avg_yearly)
    else:
        obs_year = obs_year_base  # e.g. 2025 in a Jan-2026 report
        if slide_hii_yearly:
            update_obs_vs_hii_yearly(
                slide    = slide_hii_yearly,
                obs_year = obs_year,
                # HII yearly forecast for obs_year initialised in January of that year
                path_fcst = builder.build_hii_forecast_yearly(obs_year, obs_year, 1),
                # Yearly observed vs avg30y anomaly (used as context)
                path_obs  = builder.build_diff_obs_yearly_jan_report(obs_year, year, compare_to="AVG30Y"),
                # Yearly diff: HII forecast vs observed
                path_diff = builder.build_diff_obs_yearly_jan_report(obs_year, year, compare_to="HII"),
            )
        if slide_avg_yearly:
            update_obs_vs_avg_yearly(
                slide    = slide_avg_yearly,
                obs_year = obs_year,
                path_avg  = builder.build_avg30y_yearly(),   # yearly avg30y
                path_obs  = builder.build_diff_obs_yearly_jan_report(obs_year, year, compare_to="AVG30Y"),
                path_diff = builder.build_diff_obs_yearly_jan_report(obs_year, year, compare_to="AVG30Y"),
            )

    for slide, (obs_year, obs_month) in zip(hii_slides, past_months_210):
        update_obs_vs_hii(
            slide      = slide,
            obs_year   = obs_year,
            obs_month  = obs_month,
            path_fcst  = builder.build_hii_forecast_path(obs_year, obs_month, obs_year, obs_month),
            path_obs   = builder.build_obs_path(obs_year, obs_month),
            path_diff  = builder.build_diff_obs_vs_forecast_path(obs_year, obs_month, "HII"),
            tbl_data   = build_obs_diff_table("HII", obs_year, obs_month),
        )

    for slide, (obs_year, obs_month) in zip(tmd_slides, past_months_210):
        update_obs_vs_tmd(
            slide      = slide,
            obs_year   = obs_year,
            obs_month  = obs_month,
            path_fcst  = builder.build_tmd_forecast_path(obs_year, obs_month, obs_year, obs_month),
            path_obs   = builder.build_obs_path(obs_year, obs_month),
            path_diff  = builder.build_diff_obs_vs_forecast_path(obs_year, obs_month, "TMD"),
            tbl_data   = build_obs_diff_table("TMD", obs_year, obs_month),
        )

    for slide, (obs_year, obs_month) in zip(om_slides, past_months_210):
        update_obs_vs_om(
            slide      = slide,
            obs_year   = obs_year,
            obs_month  = obs_month,
            path_fcst  = builder.build_onemap_path(obs_year, obs_month, obs_year, obs_month, model_type="MFCST"),
            path_obs   = builder.build_obs_path(obs_year, obs_month),
            path_diff  = builder.build_diff_obs_vs_forecast_path(obs_year, obs_month, "OM"),
            tbl_data   = build_obs_diff_table("OM", obs_year, obs_month),
        )

    for slide, (obs_year, obs_month) in zip(avg_slides, past_months_210):
        update_obs_vs_avg(
            slide      = slide,
            obs_year   = obs_year,
            obs_month  = obs_month,
            path_avg   = builder.build_avg30y_monthly(obs_month),
            path_obs   = builder.build_obs_path(obs_year, obs_month),
            path_diff  = builder.build_obs_path(obs_year, obs_month, is_diff=True),
        )

    # ------------------------------------------------------------------
    # Group 2.11 — Forecast vs avg30y (t1–t6, one fixed slide each)
    # ------------------------------------------------------------------
    for slot, m in zip(["t1", "t2", "t3", "t4", "t5", "t6"], target_months):
        slide = _get_slide(report, f"tag_fcst_vs_avg_{slot}")
        if slide:
            update_fcst_vs_avg(
                slide         = slide,
                target_year   = m["year"],
                target_month  = m["month"],
                path_avg      = builder.build_avg30y_monthly(m["month"], area="country"),
                path_tmd      = builder.build_tmd_forecast_path(year, month, m["year"], m["month"], area="country"),
                path_tmd_anom = builder.build_tmd_forecast_path(year, month, m["year"], m["month"], is_diff=True, area="country"),
                path_hii      = builder.build_hii_forecast_path(year, month, m["year"], m["month"], area="country"),
                path_hii_anom = builder.build_hii_forecast_path(year, month, m["year"], m["month"], is_diff=True, area="country"),
            )

    # ------------------------------------------------------------------
    # Group 2.12 — Basin-level forecast (image + table pairs)
    # ------------------------------------------------------------------
    def _build_basin_paths(builder_fn):
        """Build t1-t6 basin paths using a 1-arg callable(month_info)."""
        return [builder_fn(m) for m in target_months]

    # ── HII basin ────────────────────────────────────────────────────
    slide = _get_slide(report, "tag_hii_basin_monthly")
    if slide:
        update_hii_basin_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths_fcst = _build_basin_paths(
                lambda m: builder.build_hii_forecast_path(year, month, m["year"], m["month"], area="basin")
            ),
            paths_anom = _build_basin_paths(
                lambda m: builder.build_hii_forecast_path(year, month, m["year"], m["month"], is_diff=True, area="basin")
            ),
        )

    slide = _get_slide(report, "tag_hii_basin_tbl")
    if slide:
        update_hii_basin_tbl(slide, year, month, tbl_data=tbl_basin_hii)

    # ── One Map Mean basin ────────────────────────────────────────────
    slide = _get_slide(report, "tag_om_basin_monthly")
    if slide:
        update_om_basin_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths_fcst = _build_basin_paths(
                lambda m: builder.build_onemap_path(year, month, m["year"], m["month"], model_type="MFCST", area="basin")
            ),
            paths_anom = _build_basin_paths(
                lambda m: builder.build_onemap_path(year, month, m["year"], m["month"], model_type="MFCST", is_diff=True, area="basin")
            ),
        )

    slide = _get_slide(report, "tag_om_basin_tbl")
    if slide:
        update_om_basin_tbl(slide, year, month, tbl_data=tbl_basin_om)

    # ── One Map Upper basin ───────────────────────────────────────────
    slide = _get_slide(report, "tag_om_upper_basin_monthly")
    if slide:
        update_om_upper_basin_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths_fcst = _build_basin_paths(
                lambda m: builder.build_onemap_path(year, month, m["year"], m["month"], model_type="UFCST", area="basin")
            ),
            paths_anom = _build_basin_paths(
                lambda m: builder.build_onemap_path(year, month, m["year"], m["month"], model_type="UFCST", is_diff=True, area="basin")
            ),
        )

    slide = _get_slide(report, "tag_om_upper_basin_tbl")
    if slide:
        update_om_upper_basin_tbl(slide, year, month, tbl_data=tbl_basin_om_u)

    # ── One Map Lower basin ───────────────────────────────────────────
    slide = _get_slide(report, "tag_om_lower_basin_monthly")
    if slide:
        update_om_lower_basin_monthly(
            slide      = slide,
            init_year  = year,
            init_month = month,
            paths_fcst = _build_basin_paths(
                lambda m: builder.build_onemap_path(year, month, m["year"], m["month"], model_type="LFCST", area="basin")
            ),
            paths_anom = _build_basin_paths(
                lambda m: builder.build_onemap_path(year, month, m["year"], m["month"], model_type="LFCST", is_diff=True, area="basin")
            ),
        )

    slide = _get_slide(report, "tag_om_lower_basin_tbl")
    if slide:
        update_om_lower_basin_tbl(slide, year, month, tbl_data=tbl_basin_om_l)

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------
    output_manager = OutputManager(settings.paths.output_dir)
    output_path    = output_manager.build_output_path(OutputSpec(year=year, month=month))
    report.save_report(output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="HII 6-month rainfall forecast report generator")
    parser.add_argument("-y", "--year",  type=int, required=True, help="Report year (AD)")
    parser.add_argument("-m", "--month", type=int, required=True, help="Report month (1-12)")
    args = parser.parse_args()
    main(args.year, args.month)
