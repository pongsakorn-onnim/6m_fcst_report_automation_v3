from pathlib import Path
from src.core.rain_data_service import RainDataService

excel_path = Path(
    r"D:\HII\extract_rain_to_excel\outputs\extract\rain_summary_202602.xlsx"
)

service = RainDataService(excel_path)

table = service.build_table(zone_type="Region", model="HII")

print("Months:", table["months"])
print("First row:", table["rows"][0])