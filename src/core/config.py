"""
Application Configuration Loader
--------------------------------
- โหลด config.yaml
- map เข้า dataclass
- แปลง path เป็น pathlib.Path
- ไม่ใส่ business logic
"""

from __future__ import annotations

import logging
import yaml
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


# ==========================================================
# Dataclasses
# ==========================================================

@dataclass
class PathsConfig:
    # Shared resources
    shared_dir: Path = Path()
    analog_years_csv: Path = Path()

    # External rain extraction project
    spatial_rain_extract_dir: Path = Path()

    # Internal project paths
    templates_dir: Path = Path("templates")
    output_dir: Path = Path("output")


@dataclass
class UrlsConfig:
    hii_yearly_base: str = ""
    hii_monthly_base: str = ""
    fcst_hii_diff_year_base: str = ""
    fcst_onemap_tmd_base: str = ""
    avg30y_base: str = ""
    avg30y_yearly_image: str = ""
    diff_obs_hii_base: str = ""


# ==========================================================
# Main AppConfig
# ==========================================================

class AppConfig:
    """
    Central configuration object.
    Responsible ONLY for:
    - Loading YAML
    - Mapping into dataclasses
    - Normalizing paths
    """

    def __init__(self, config_filename: str = "config.yaml"):

        # Root project directory
        self.base_dir: Path = Path(__file__).resolve().parents[2]

        full_config_path = self.base_dir / config_filename

        config_dict: dict = {}

        try:
            with full_config_path.open("r", encoding="utf-8") as f:
                parsed_yaml = yaml.safe_load(f)
                if parsed_yaml:
                    config_dict = parsed_yaml
        except FileNotFoundError:
            logger.warning(
                "Config file not found at %s. Using defaults.",
                full_config_path
            )
        except yaml.YAMLError as e:
            logger.error(
                "YAML parsing error at %s: %s",
                full_config_path,
                e
            )

        # Map into dataclasses
        self.paths = self._load_paths(config_dict.get("paths", {}))
        self.urls = UrlsConfig(**config_dict.get("urls", {}))

    # ------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------

    def _load_paths(self, paths_dict: dict) -> PathsConfig:
        """
        Convert string paths from YAML into pathlib.Path.
        """

        def to_path(value: Optional[str], relative_to_base: bool = False) -> Path:
            if not value:
                return Path()

            p = Path(value)

            # If relative path (e.g. "templates"), resolve from project root
            if relative_to_base and not p.is_absolute():
                return self.base_dir / p

            return p

        return PathsConfig(
            shared_dir=to_path(paths_dict.get("shared_dir")),
            analog_years_csv=to_path(paths_dict.get("analog_years_csv")),
            spatial_rain_extract_dir=to_path(
                paths_dict.get("spatial_rain_extract_dir")
            ),
            templates_dir=to_path(
                paths_dict.get("templates_dir", "templates"),
                relative_to_base=True,
            ),
            output_dir=to_path(
                paths_dict.get("output_dir", "output"),
                relative_to_base=True,
            ),
        )


# ==========================================================
# Singleton Instance
# ==========================================================

settings = AppConfig()