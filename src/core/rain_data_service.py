"""
Rain Data Service
-----------------
Builds table-ready data structures for PPT from pre-processed CSV/Excel files.

Data sources (primary — senior-provided files):
  Forecast tables (2.7, 2.8, 2.9, 2.12):
    OM_W/OM_U/OM_L  → YYYYMM_om_{region|basin}.csv  + YYYYMM_diff_{region|basin}.csv
    HII             → analog_years.csv + monthlyrain_{region|basin}.csv

  Obs-vs-fcst tables (2.10):
    HII             → HIIObserve_forecast_region_{year}.xlsx
    TMD             → TMDObserve_forecast_region_{year}.xlsx
    OM              → Observe_OMWforecast_{year}.xlsx

All paths come from rain_services.config (shared_rain_services package).
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Literal

import pandas as pd
from rain_services import config as rain_config

logger = logging.getLogger(__name__)

ZoneType = Literal["Region", "Basin"]

_REGION_NAME_MAP = {
    "ภาคเหนือ":               "เหนือ",
    "ภาคตะวันออกเฉียงเหนือ":  "ตะวันออกเฉียงเหนือ",
    "ภาคกลาง":                "กลาง",
    "ภาคตะวันออก":            "ตะวันออก",
    "ภาคใต้ฝั่งตะวันออก":    "ใต้ฝั่งตะวันออก",
    "ภาคใต้ฝั่งตะวันตก":     "ใต้ฝั่งตะวันตก",
}

THAI_MONTHS = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.",
    5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.",
    9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}

# OM model → candidate column names in the CSV (inconsistent across months: MEAN vs WEIGHT)
_OM_MODEL_COLS = {"OM_W": ("MEAN", "WEIGHT"), "OM_U": ("UPPER",), "OM_L": ("LOWER",)}


def _resolve_om_col(model: str, df_columns) -> str:
    """Return the first matching column name for the given OM model."""
    for col in _OM_MODEL_COLS[model]:
        if col in df_columns:
            return col
    raise KeyError(
        f"OM model '{model}': none of {_OM_MODEL_COLS[model]} found in columns {list(df_columns)}"
    )

# Obs-vs-fcst region Excel filename pattern per model
_OBS_DIFF_REGION_FILES = {
    "HII": lambda y: f"HIIObserve_forecast_region_{y}.xlsx",
    "TMD": lambda y: f"TMDObserve_forecast_region_{y}.xlsx",
    "OM_W": lambda y: f"Observe_OMWforecast_{y}.xlsx",
}


class RainDataService:
    """
    Loads and serves rainfall table data for a given report (year, month).

    Usage:
        svc = RainDataService(2026, 3)
        tbl = svc.build_table("Region", "OM_W")
        tbl = svc.build_table("Basin",  "HII")
    """

    def __init__(self, year: int, month: int):
        self.year  = year
        self.month = month
        self._cache: dict = {}

    # ------------------------------------------------------------------
    # Cache helpers
    # ------------------------------------------------------------------

    def _cached(self, key: str, loader):
        if key not in self._cache:
            self._cache[key] = loader()
        return self._cache[key]

    def _get_avg30y_region(self) -> pd.DataFrame:
        def load():
            df = pd.read_csv(rain_config.AVG30Y_REGION_CSV)
            return df[["MONTH", "REG_CODE", "MEAN_OBS"]]
        return self._cached("avg30y_region", load)

    def _get_avg30y_basin(self) -> pd.DataFrame:
        """Load avg30y basin CSV — drops '-is' island rows, casts MB_CODE to int."""
        def load():
            df = pd.read_csv(rain_config.AVG30Y_BASIN_CSV)
            df = df[pd.to_numeric(df["MB_CODE"], errors="coerce").notna()].copy()
            df["MB_CODE"] = df["MB_CODE"].astype(int)
            return df[["MONTH", "MB_CODE", "MEAN"]]
        return self._cached("avg30y_basin", load)

    def _get_obs_region(self) -> pd.DataFrame:
        def load():
            return pd.read_csv(rain_config.MONTHLY_RAIN_REGION_CSV)
        return self._cached("obs_region", load)

    def _get_obs_basin(self) -> pd.DataFrame:
        """Load monthly observed basin CSV — drops '-is' island rows, casts MB_CODE to int."""
        def load():
            df = pd.read_csv(rain_config.MONTHLY_RAIN_BASIN_CSV)
            df = df[pd.to_numeric(df["MB_CODE"], errors="coerce").notna()].copy()
            df["MB_CODE"] = df["MB_CODE"].astype(int)
            return df
        return self._cached("obs_basin", load)

    def _get_analog_year(self) -> int:
        def load():
            df = pd.read_csv(rain_config.ANALOG_YEARS_CSV_PATH)
            row = df[(df["target_year"] == self.year) & (df["init_month"] == self.month)]
            if row.empty:
                raise ValueError(
                    f"No analog year entry for target_year={self.year}, init_month={self.month}. "
                    f"Update {rain_config.ANALOG_YEARS_CSV_PATH}"
                )
            return int(row["analog_year"].iloc[0])
        return self._cached("analog_year", load)

    def _get_target_months(self) -> list:
        """Returns list of 6 dicts {year, month} for t1-t6."""
        def load():
            from src.core.ppt_tools.text_handler import get_next_months
            return get_next_months(self.year, self.month, 6)
        return self._cached("target_months", load)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def build_table(self, zone_type: ZoneType, model: str) -> dict | None:
        """
        Build a PPT-ready table dict for the given zone type and model.
        Tries senior's CSV files first; falls back to extracted Excel if missing.

        Returns None (with a warning) if both sources fail.

        Args:
            zone_type: "Region" or "Basin"
            model:     "HII", "OM_W", "OM_U", or "OM_L"

        Returns:
            {
              "zone_type": str,
              "model":     str,
              "months":    [Thai abbrev, ...],   # 6 entries, metadata only
              "rows": [
                  {"code": int, "name": str,
                   "values": [{"anomaly": float, "percent": float}, ...]},
                  ...
              ]
            }
        """
        try:
            if model == "HII":
                return self._build_hii_table(zone_type)
            elif model in _OM_MODEL_COLS:
                return self._build_om_table(zone_type, model)
            else:
                logger.warning(f"build_table: unknown model '{model}' — skipped.")
                return None
        except Exception as e:
            logger.warning(f"build_table({zone_type!r}, {model!r}) primary failed: {e} — trying fallback.")
            return self._build_from_extracted_excel(zone_type, model)

    def _build_from_extracted_excel(self, zone_type: ZoneType, model: str) -> dict | None:
        """Fallback: read forecast table from extract_rain_to_excel output Excel."""
        yyyymm = f"{self.year}{self.month:02d}"
        path = rain_config.EXTRACT_RAIN_EXCEL_DIR / f"rain_summary_{yyyymm}.xlsx"

        if not path.exists():
            logger.warning(f"Fallback Excel not found: {path}")
            return None

        try:
            df = pd.read_excel(path, sheet_name=zone_type)
            df = df[df["model"] == model].copy()
            if df.empty:
                logger.warning(f"Fallback Excel: no data for model={model!r}, zone={zone_type!r} in {path.name}")
                return None

            code_col = "REG_CODE" if zone_type == "Region" else "MB_CODE"
            name_col = "FIRST_REGI" if zone_type == "Region" else "MBASIN_T"

            df = df.sort_values("lead_time")
            sorted_leads = sorted(df["lead_time"].unique())

            month_labels = []
            for lead in sorted_leads:
                tm = df[df["lead_time"] == lead]["target_month"].iloc[0]  # "YYYY-MM"
                month_labels.append(THAI_MONTHS[int(str(tm).split("-")[1])])

            rows_by_code: dict = {}
            for lead in sorted_leads:
                for _, row in df[df["lead_time"] == lead].iterrows():
                    code = int(row[code_col])
                    raw_name = str(row[name_col]).strip()
                    name = _REGION_NAME_MAP.get(raw_name, raw_name) if zone_type == "Region" else raw_name
                    if code not in rows_by_code:
                        rows_by_code[code] = {"code": code, "name": name, "values": []}
                    rows_by_code[code]["values"].append({
                        "anomaly": float(row["anomaly"]),
                        "percent": float(row["percent_anomaly"]),
                    })

            rows = [rows_by_code[c] for c in sorted(rows_by_code)]
            logger.info(f"Fallback Excel used for build_table({zone_type!r}, {model!r})")
            return {"zone_type": zone_type, "model": model, "months": month_labels, "rows": rows}

        except Exception as e:
            logger.warning(f"Fallback Excel read failed ({zone_type!r}, {model!r}): {e}")
            return None

    # ------------------------------------------------------------------
    # HII forecast table (analog year + observed monthly data)
    # ------------------------------------------------------------------

    def _build_hii_table(self, zone_type: ZoneType) -> dict:
        analog_base = self._get_analog_year()
        targets     = self._get_target_months()

        if zone_type == "Region":
            obs_df   = self._get_obs_region()
            code_col = "REG_CODE"
            name_col = "FIRST_REGI"
        else:
            obs_df   = self._get_obs_basin()
            code_col = "MB_CODE"
            name_col = "MBASIN_T"

        month_labels = [THAI_MONTHS[m["month"]] for m in targets]
        rows_by_code: dict = {}

        for t in targets:
            analog_year = analog_base + (t["year"] - self.year)
            subset = obs_df[
                (obs_df["YEAR"] == analog_year) & (obs_df["MONTH"] == t["month"])
            ]
            if subset.empty:
                logger.warning(
                    f"HII {zone_type}: no observed data for analog year {analog_year}, "
                    f"month {t['month']}"
                )
                continue

            for _, row in subset.iterrows():
                code = int(row[code_col])
                raw_name = str(row[name_col]).strip()
                name = _REGION_NAME_MAP.get(raw_name, raw_name) if zone_type == "Region" else raw_name

                if code not in rows_by_code:
                    rows_by_code[code] = {"code": code, "name": name, "values": []}

                rows_by_code[code]["values"].append({
                    "anomaly": float(row["MEAN_DIFF"]),
                    "percent": float(row["PERCENTAGE_DIFF"]),
                })

        rows = [rows_by_code[c] for c in sorted(rows_by_code)]
        return {"zone_type": zone_type, "model": "HII", "months": month_labels, "rows": rows}

    # ------------------------------------------------------------------
    # OM forecast table (OM_W / OM_U / OM_L via MEAN/UPPER/LOWER columns)
    # ------------------------------------------------------------------

    def _build_om_table(self, zone_type: ZoneType, model: str) -> dict:
        yyyymm = f"{self.year}{self.month:02d}"

        if zone_type == "Region":
            diff_path    = rain_config.ONEMAP_REGION_CSV_DIR / f"{yyyymm}_diff_region.csv"
            avg_df       = self._get_avg30y_region()
            code_col     = "REG_CODE"
            name_col     = "REG_T"
            avg_code_col = "REG_CODE"
            avg_val_col  = "MEAN_OBS"
        else:
            diff_path    = rain_config.ONEMAP_BASIN_CSV_DIR / f"{yyyymm}_diff_basin.csv"
            avg_df       = self._get_avg30y_basin()
            code_col     = "BASIN_CODE"
            name_col     = "BASIN_T"
            avg_code_col = "MB_CODE"
            avg_val_col  = "MEAN"

        if not diff_path.exists():
            raise FileNotFoundError(f"OM diff CSV not found: {diff_path}")

        diff_df      = pd.read_csv(diff_path)
        model_col    = _resolve_om_col(model, diff_df.columns)
        targets      = self._get_target_months()
        month_labels = [THAI_MONTHS[m["month"]] for m in targets]
        rows_by_code: dict = {}

        for t in targets:
            sub_diff = diff_df[diff_df["MONTH"] == t["month"]]
            sub_avg  = avg_df[avg_df["MONTH"]   == t["month"]]

            for _, row in sub_diff.iterrows():
                code    = int(row[code_col])
                anomaly = float(row[model_col])

                avg_row = sub_avg[sub_avg[avg_code_col] == code]
                if avg_row.empty:
                    logger.warning(
                        f"OM {zone_type} {model}: no avg30y for code={code}, month={t['month']}"
                    )
                    pct = float("nan")
                else:
                    avg_val = float(avg_row[avg_val_col].iloc[0])
                    pct = (anomaly / avg_val * 100) if avg_val != 0 else float("nan")

                raw_name = str(row[name_col]).strip()
                name = _REGION_NAME_MAP.get(raw_name, raw_name) if zone_type == "Region" else raw_name

                if code not in rows_by_code:
                    rows_by_code[code] = {"code": code, "name": name, "values": []}
                rows_by_code[code]["values"].append({"anomaly": anomaly, "percent": pct})

        rows = [rows_by_code[c] for c in sorted(rows_by_code)]
        return {"zone_type": zone_type, "model": model, "months": month_labels, "rows": rows}


# ======================================================================
# Obs-vs-Forecast table (Group 2.10) — module-level, region only
# ======================================================================

def build_obs_diff_table(
    model: str, year: int, month: int,
    init_year: int | None = None, init_month: int | None = None,
) -> dict:
    """
    Load observed-vs-forecast diff data for a single past month (Group 2.10).
    Tries senior's Excel first; falls back to extracted Excel if missing.

    Args:
        model:       "HII", "TMD", or "OM_W"
        year:        Observed year
        month:       Observed month (1–12)
        init_year:   Report init year  (required for fallback file lookup)
        init_month:  Report init month (required for fallback file lookup)

    Returns:
        {thai_region_short_name: {"anomaly": float, "percent": float}}
        Returns {} if both sources fail.
    """
    filename_fn = _OBS_DIFF_REGION_FILES.get(model)
    if filename_fn is None:
        logger.warning(f"build_obs_diff_table: unknown model '{model}'")
        return {}

    path = rain_config.DIFF_REGION_EXCEL_DIR / filename_fn(year)

    if path.exists():
        try:
            xl = pd.ExcelFile(path)
            df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
            df = df[(df["YEAR"] == year) & (df["MONTH"] == month)]

            if df.empty:
                logger.warning(
                    f"Obs-diff ({model}): no rows for {year}-{month:02d} in {path.name}"
                )
            else:
                # Normalize diff column — OMW region uses 'obs_anom_fcst'; HII/TMD use 'obs_fcst'
                diff_col = "obs_anom_fcst" if "obs_anom_fcst" in df.columns else "obs_fcst"
                result = {}
                for _, row in df.iterrows():
                    raw_name = str(row["FIRST_REGI"]).strip()
                    name = _REGION_NAME_MAP.get(raw_name, raw_name)
                    result[name] = {
                        "anomaly": float(row[diff_col]),
                        "percent": float(row["anom_per"]),
                    }
                return result

        except Exception as e:
            logger.warning(f"Obs-diff primary read failed ({model}, {year}-{month:02d}): {e}")
    else:
        logger.warning(f"Obs-diff file not found: {path}")

    # Fallback
    if init_year is None or init_month is None:
        logger.warning(f"Obs-diff fallback skipped: init_year/init_month not provided.")
        return {}
    return _build_obs_diff_from_extracted(model, year, month, init_year, init_month)


def _build_obs_diff_from_extracted(
    model: str, year: int, month: int, init_year: int, init_month: int
) -> dict:
    """Fallback: read obs-diff data from extract_rain_to_excel output Excel."""
    init_yyyymm = f"{init_year}{init_month:02d}"
    path = rain_config.EXTRACT_RAIN_EXCEL_DIR / f"obs_diff_summary_{init_yyyymm}.xlsx"

    if not path.exists():
        logger.warning(f"Fallback obs-diff Excel not found: {path}")
        return {}

    try:
        df = pd.read_excel(path, sheet_name="ObsDiff_Region")
        df = df[(df["model"] == model) & (df["obs_year"] == year) & (df["obs_month"] == month)]
        if df.empty:
            logger.warning(
                f"Fallback obs-diff: no data for model={model!r}, {year}-{month:02d} in {path.name}"
            )
            return {}

        result = {}
        for _, row in df.iterrows():
            raw_name = str(row["FIRST_REGI"]).strip()
            name = _REGION_NAME_MAP.get(raw_name, raw_name)
            result[name] = {
                "anomaly": float(row["anomaly"]),
                "percent": float(row["percent_anomaly"]),
            }
        logger.info(f"Fallback Excel used for obs-diff ({model}, {year}-{month:02d})")
        return result

    except Exception as e:
        logger.warning(f"Fallback obs-diff Excel read failed ({model}, {year}-{month:02d}): {e}")
        return {}
