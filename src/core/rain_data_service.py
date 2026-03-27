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


logger = logging.getLogger(__name__)

ZoneType = Literal["Region", "Basin"]


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

        self.excel_path = excel_path

    # ----------------------------------------------------------
    # Internal
    # ----------------------------------------------------------

    def _load_sheet(self, zone_type: ZoneType) -> pd.DataFrame:
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
        else:
            code_col = "MB_CODE"
            name_col = "MBASIN_N"

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