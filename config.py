import os

BOT_TOKEN = os.environ.get("BOT_TOKEN", "sizning_tokeningiz")
ADMIN_IDS_ENV = os.environ.get("ADMIN_ID", "")
if ADMIN_IDS_ENV:
    ADMIN_ID = [int(x.strip()) for x in ADMIN_IDS_ENV.split(",")]
else:
    ADMIN_ID = [5171707160, 987654321]  # lokal uchun o'z ID laringiz

COMPANIES_FILE = "firmalar.json"
TEMPLATES_FOLDER = "shablonlar"
DOCS_FOLDER = "markitovkalar"