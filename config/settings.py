# All hardcoded configuration values for Feedback Automation Tool

LOB_OPTIONS = ["Select", "Tech Certs", "SEPO", "OC,DD,BC", "All"]

# Input sheet column names (sheet: "Automation")
INPUT_COL_SR = "Sr No"
INPUT_COL_DATE = "Date"
INPUT_COL_BEST_PART = "What was the best part of this session and how has it helped you?"
INPUT_COL_RATING = "Rate the overall satisfaction level of this session"
INPUT_COL_IMPROVEMENT = "What according to you could be improved in this session?"
INPUT_COL_PL = "PL Name"
INPUT_COL_COURSE = "Course Name"
INPUT_COL_TOPIC = "Topic Name"
INPUT_COL_LOB = "LOB"

# Excel report column headers (row 3) — 4 columns: A=Sr No, B=Best Part, C=Rating, D=Improvement
EXCEL_HEADERS = [
    "Sr No",
    "What was the best part of this session and how has it helped you?",
    "Rate the overall satisfaction level of this session",
    "What according to you could be improved in this session?",
]

# Styling
GREEN_HEX = "00B050"
FOOTER_TEXT = "Developed by EMERITUS — Feedback Automation Tool"

# Column widths (openpyxl units) — A=Sr No, B=Best Part, C=Rating, D=Improvement
COL_WIDTHS = {"A": 18.86, "B": 42.86, "C": 42.57, "D": 47.86}
ZOOM_SCALE = 90

# Output directories (relative to app working directory)
EXCEL_OUTPUT_DIR = "output/excel"
PDF_OUTPUT_DIR   = "output/pdf"

# Airtable
AIRTABLE_BASE_ID = "app7HKMKVF1SlHn9e"
AIRTABLE_TABLE_NAME = "Live Sessions - India/APAC"

# Allowed users (email -> display name)
ALLOWED_USERS = {
    "hariharan.v@emeritus.org": "Hariharan",
    "mohammed.shakeel@emeritus.org": "Shakeel",
}
