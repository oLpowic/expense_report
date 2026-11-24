from pathlib import Path
import pandas as pd
import datetime

REPO_PATH = Path(__file__).resolve().parents[1]   # .../expense_report
BANK_STATEMENTS_DIR = REPO_PATH / "bank_statements"

if __name__ == "__main__":
    print(REPO_PATH)
    files = BANK_STATEMENTS_DIR.glob("*.xlsx")
    df = pd.concat((pd.read_excel(f) for f in files), ignore_index=True)
    print(df)
    # x = datetime.datetime.now()
    # print(x)
    
    