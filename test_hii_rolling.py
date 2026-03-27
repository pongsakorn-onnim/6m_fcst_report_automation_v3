from src.core.config import settings
from src.core.analog_year_service import AnalogYearService
from src.core.url_builder import UrlBuilder

# 1. init services
analog_service = AnalogYearService(settings.paths.analog_years_csv)

builder = UrlBuilder(**settings.urls.__dict__)

# 2. parameters
target_year_be = 2566
init_year_ce = 2023
init_month = 10

# 3. get analog year
analog_year_ce = analog_service.get_analog_year_ce(
    target_year_be,
    init_month
)

print("Analog CE:", analog_year_ce)

# 4. generate URLs
for lead in range(6):
    url = builder.forecast_monthly(
        model="HII",
        init_year=init_year_ce,
        init_month=init_month,
        lead=lead,
        analog_year=analog_year_ce,
    )
    print(f"Lead {lead}:", url)