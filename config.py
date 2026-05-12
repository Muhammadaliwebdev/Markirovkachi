import os

BOT_TOKEN     = os.environ.get("BOT_TOKEN", "")
ADMIN_ID      = int(os.environ.get("ADMIN_ID", "0"))
COMPANIES_FILE  = "firmalar.json"
TEMPLATES_FOLDER = "shablonlar"
DOCS_FOLDER    = "markitovkalar"