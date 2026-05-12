import os

BOT_TOKEN     = os.environ.get("BOT_TOKEN", "")
ADMIN_ID = [int(x) for x in os.environ.get("ADMIN_ID", "0").split(",")]
COMPANIES_FILE  = "firmalar.json"
TEMPLATES_FOLDER = "shablonlar"
DOCS_FOLDER    = "markitovkalar"