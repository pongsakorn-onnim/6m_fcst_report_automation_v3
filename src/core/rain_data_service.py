"""
Rain Data Service
-----------------
Responsible for:
- Loading Region / Basin sheet
- Filtering by model
- Preparing table-ready structure for PPT
- Mapping target_month -> Thai month label
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Literal
import pandas as pd
from rain_services import config as rain_config


logger = logging.getLogger(__name__)

ZoneType = Literal["Region", "Basin"]

# FIRST_REGI in the shapefile uses full names with ภาค prefix,
# but the template table rows use the short form without it.
_REGION_NAME_MAP = {
    "ภาคเหนือ":               "เหนือ",
    "ภาคตะวันออกเฉียงเหนือ":  "ตะวันออกเฉียงเหนือ",
    "ภาคกลาง":                "กลาง",
    "ภาคตะวันออก":            "ตะวันออก",
    "ภาคใต้ฝั่งตะวันออก":    "ใต้ฝั่งตะวันออก",
    "ภาคใต้ฝั่งตะวันตก":     "ใต้ฝั่งตะวันตก",
}

THAI_MONTHS = {
    1: "ม.ค.",
    2: "ก.พ.",
    3: "มี.ค.",
    4: "เม.ย.",
    5: "พ.ค.",
    6: "มิ.ย.",
    7: "ก.ค.",
    8: "ส.ค.",
    9: "ก.ย.",
    10: "ต.ค.",
    11: "พ.ย.",
    12: "ธ.ค.",
}


class RainDataService:

    def __init__(self, excel_path: Path):
        if not excel_path.exists():
            raise FileNotFoundError(f"Rain summary file not found: {excel_path}")
        self.excel_path   = excel_path
        self._dataframes  = None

    @classmethod
    def from_dataframes(cls, dataframes: dict) -> "RainDataService":
        """Create from pre-extracted DataFrames (Option B — no Excel file needed)."""
        instance = cls.__new__(cls)
        instance.excel_path  = None
        instance._dataframes = dataframes
        return instance

    # ----------------------------------------------------------
    # Internal
    # ----------------------------------------------------------

    def _load_sheet(self, zone_type: ZoneType) -> pd.DataFrame:
        if self._dataframes is not None:
            df = self._dataframes.get(zone_type, pd.DataFrame()).copy()
        else:
            df = pd.read_excel(self.excel_path, sheet_name=zone_type)

        required_cols = {
            "model",
            "lead_time",
            "target_month",
            "anomaly",
            "percent_anomaly",
        }

        if not required_cols.issubset(df.columns):
            raise ValueError(f"Missing required columns in {zone_type} sheet")

        return df

    @staticmethod
    def _month_label(target_month: str) -> str:
        # target_month format: 2026-02
        month = int(target_month.split("-")[1])
        return THAI_MONTHS[month]

    # ----------------------------------------------------------
    # Public API
    # ----------------------------------------------------------

    def build_table(
        self,
        zone_type: ZoneType,
        model: str,
    ) -> dict:
        """
        Returns table-ready structure for PPT.
        """

        df = self._load_sheet(zone_type)

        # filter model
        df = df[df["model"] == model]

        # ensure lead 0-5 only
        df = df[df["lead_time"].between(0, 5)]

        if df.empty:
            raise ValueError(f"No data found for model={model} in {zone_type}")

        # sort properly
        df = df.sort_values(["lead_time"])

        # determine row id column
        if zone_type == "Region":
            code_col = "REG_CODE"
            name_col = "FIRST_REGI"
            df[name_col] = df[name_col].map(lambda n: _REGION_NAME_MAP.get(n, n))
        else:
            code_col = "MB_CODE"
            name_col = "MBASIN_T"

        # month headers (unique ordered by lead)
        month_labels = (
            df.sort_values("lead_time")
            .drop_duplicates("lead_time")
            .apply(lambda r: self._month_label(r["target_month"]), axis=1)
            .tolist()
        )

        rows = []

        for code in sorted(df[code_col].unique()):
            sub = df[df[code_col] == code].sort_values("lead_time")

            values = []

            for _, r in sub.iterrows():
                values.append(
                    {
                        "anomaly": float(r["anomaly"]),
                        "percent": float(r["percent_anomaly"]),
                    }
                )

            rows.append(
                {
                    "code": code,
                    "name": sub[name_col].iloc[0],
                    "values": values,
                }
            )

        return {
            "zone_type": zone_type,
            "model": model,
            "months": month_labels,
            "rows": rows,
        }


def build_obs_diff_table(model: str, year: int, month: int) -> dict:
    """
    Load observed-vs-forecast diff data for Group 2.10 directly from
    pre-extracted Excel files. Independent of RainDataService / raster pipeline.

    Args:
        model: "HII", "TMD", or "OM"
        year:  Observed year
        month: Observed month (1–12)

    Returns:
        {thai_region_name: {"anomaly": float, "percent": float}}
    """
    filename = f"{model}Observe_forecast_region_{year}.xlsx"
    path = rain_config.DIFF_REGION_EXCEL_DIR / filename

    if not path.exists():
        logger.warning(f"Obs-diff Excel not found: {path}")
        return {}

    try:
        df = pd.read_excel(path, sheet_name="Sheet 1")
    except Exception as e:
        logger.error(f"Failed to read {path}: {e}")
        return {}

    df = df[(df["YEAR"] == year) & (df["MONTH"] == month)]
    df = df.drop_duplicates(subset=["REG_CODE"], keep="first")

    result = {}
    for _, row in df.iterrows():
        raw_name = str(row["FIRST_REGI"]).strip()
        name = _REGION_NAME_MAP.get(raw_name, raw_name.removeprefix("ภาค"))
        result[name] = {
            "anomaly": float(row["obs_fcst"]),
            "percent": float(row["anom_per"]),
        }

    return result
