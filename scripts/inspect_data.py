
from pathlib import Path
import pandas as pd

project_dir = Path(__file__).resolve().parent.parent
file_path = project_dir / "data" / "clinical_study_raw.xlsx"

for sheet in ["Participants", "Visits", "Laboratory", "Adverse_Events"]:
    df = pd.read_excel(file_path, sheet_name=sheet)

    print("\n" + "=" * 70)
    print(f"SHEET: {sheet}")
    print("Columns:", list(df.columns))
    print("\nFirst 5 records:")
    print(df.head().to_string(index=False))

    if sheet == "Visits":
        print("\nVisit dates for CT001:")
        print(
            df[df["Subject_ID"] == "CT001"]
            [["Subject_ID", "Visit_Name", "Visit_Date"]]
            .to_string(index=False)
        )
