from pathlib import Path
import pandas as pd
import logging
import os
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter(
     "(%(asctime)s) | %(name)s | %(levelname)s => '%(message)s'"
)

logs_dir = Path("logs")
logs_dir.mkdir(exist_ok=True)
log_file = logs_dir / "automation.log"

file_handler = logging.FileHandler(log_file)
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

if not logger.hasHandlers():
    logger.addHandler(file_handler)

INPUT_DIR = Path('input')
OUTPUT_DIR = Path('output')
OUTPUT_FILE = OUTPUT_DIR / "master_report.xlsx"

REQUIRED_COLUMNS = ["name", "department", "amount", "date"]


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip().lower() for c in df.columns]
    return df

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_headers(df)

    # Remove duplicated columns created by messy headers
    df = df.loc[:, ~df.columns.duplicated()]

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            logger.warning(f"Missing column '{col}', creating it...")
            df[col] = None
    
    df = df[REQUIRED_COLUMNS]
    
    df = df.dropna(axis=1, how="all")

    df["name"] = df["name"].fillna("Unknown")
    df["department"] = df["department"].fillna("Unknown")
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    before = len(df)
    df = df.drop_duplicates()
    logger.info(f"Removed {before - len(df)} duplicates rows")

    return df

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not INPUT_DIR.exists():
        logger.error("Input directory does not exist.")
        return
    
    all_dfs = []

    logger.info("Scanning input directory...")

    for file in INPUT_DIR.iterdir():
        try:
            if file.suffix == ".csv":
                logger.info(f"Loading {file.name}")
                df = pd.read_csv(file)
                all_dfs.append(df)

            elif file.suffix == ".xlsx":
                logger.info(f"Loading {file.name}")
                df = pd.read_excel(file)
                all_dfs.append(df)

            else:
                logger.warning(f"Skipping unsupported file: {file.name}")
        
        except Exception as e:
            logger.error(f"Failed to load {file.name}: {e}")
        
    if not all_dfs:
        logger.error("No valid data could be loaded.")
        return

    merged_df = pd.concat(all_dfs, ignore_index=True)
    logger.info(f"Raw rows: {len(merged_df)}")

    cleaned_df = clean_dataframe(merged_df)
    cleaned_df = cleaned_df.sort_values("date")
    logger.info(f"Clean rows: {len(cleaned_df)}")

    summary = (
        cleaned_df.groupby("department")["amount"]
        .sum()
        .reset_index()
        .sort_values("amount", ascending=False)
    )

    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
            cleaned_df.to_excel(writer, index=False, sheet_name="Cleaned Data")
            summary.to_excel(writer, index=False, sheet_name="Summary")

        logger.info(f"Report generated: {OUTPUT_FILE.resolve()}")

    except Exception as e:
        logger.critical(f"Failed to write Excel report: {e}")

def send_email():

    load_dotenv()

    EMAIL_ADDRESS = os.getenv('SENDER')
    EMAIL_PASSWORD = os.getenv('PASSWORD')
    RECIEVER_ADDRESS = os.getenv('RECIEVER')

    msg = EmailMessage()
    msg['Subject'] = 'Your Master File is Ready Mr.LEKBIR'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = RECIEVER_ADDRESS
    msg.set_content('Sir, This is your excel file Cleaned , Organized , and Ready for analyst!')

    file = OUTPUT_FILE
    with open(file, "rb") as f:
        file_data = f.read()
        file_name = f.name
    
    msg.add_attachment(file_data,
                        maintype='application',
                        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        filename=file_name)
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        try:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            logger.info("Logged in...")

            smtp.send_message(msg)
            logger.info("Email has been sent!")
        except Exception as e:
            logger.error(f'Unable to sign in => {e}')

if __name__ == '__main__':
    main()
    send_email()