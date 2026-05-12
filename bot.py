import os
import io
import json
import logging
import re
import zipfile
from pathlib import Path

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)
from config import BOT_TOKEN, DOCS_FOLDER, COMPANIES_FILE, TEMPLATES_FOLDER
from config import BOT_TOKEN, DOCS_FOLDER, COMPANIES_FILE, TEMPLATES_FOLDER, ADMIN_ID

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ── Conversation states ──────────────────────────────────────────────────────
(
    ADD_EXPORTER_NAME,
    ADD_EXPORTER_ADDR,
    ADD_IMPORTER_NAME,
    ADD_IMPORTER_ADDR,
) = range(4)

# ── Tovarlar ro'yxati ────────────────────────────────────────────────────────
PRODUCTS = [
    "🧅 Лук репчатый свежий",
    "🍑 Абрикосы свежие",
    "🍒 Черешня свежая",
    "🧄 Чеснок свежий",
    "🍎 Гранат свежий",
    "🍑 Сливы свежие",
    "🍑 Персики свежие",
    "🍑 Нектарины свежие",
    "🍈 Дыни свежие",
    "🍇 Виноград свежий столовых сортов",
    "🍏 Яблоки свежие",
    "🍐 Айва свежая",
    "🥒 Огурцы свежие",
    "🍅 Томаты свежие некалиброванный",
    "🍋 Лимоны свежие некалиброванный",
    "🥬 Капуста белокочанная свежая",
    "🥬 Капуста краснокочанная свежая",
    "🌶 Перец горький свежий",
    "🫑 Перец стручковый сладкий свежий",
    "🥦 Капуста цветная свежая",
    "🥦 Капуста брокколи свежая",
    "🥬 Капуста Пекинская свежая",
    "🫚 Свекла столовая свежая",
    "🍆 Баклажаны свежие",
    "🌿 Дайкон свежий",
    "🥕 Морковь свежая",
    "🌱 Редька свежая",
    "🌿 Репа свежая",
    "🌿 Зелень Укропа свежая",
    "🌿 Зелень Петрушки свежая",
    "🌿 Зелень кинзы свежая",
    "🧅 Лук зеленый свежий",
    "🔴 Редис свежий",
    "🥗 Салат Руккола листовой свежий",
]
# ── Firmalar JSON ─────────────────────────────────────────────────────────────

def load_companies():
    path = Path(COMPANIES_FILE)
    if not path.exists():
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save_companies(companies):
    path = Path(COMPANIES_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(companies, f, ensure_ascii=False, indent=2)


def get_company_by_id(company_id):
    for c in load_companies():
        if c["id"] == company_id:
            return c
    return None


# ── DOCX shablon almashtirish ────────────────────────────────────────────────

def _replace_split_runs(xml, old_text, new_text):
    """
    Word bitta so'zni bir nechta <w:r><w:t> ga bo'lganda topib almashtiradi.
    Birinchi runda yangi matn qo'yiladi, qolganlarining matni bo'shatiladi.
    """
    pattern = r'(<w:r\b[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>)(.*?)(</w:t></w:r>)'
    runs = list(re.finditer(pattern, xml, re.DOTALL))
    if not runs:
        return xml

    texts = [m.group(2) for m in runs]
    combined = "".join(texts)

    if old_text not in combined:
        return xml

    start_char = combined.index(old_text)
    end_char = start_char + len(old_text)

    char_pos = 0
    first_run = None
    last_run = None
    first_run_char_offset = 0

    for i, t in enumerate(texts):
        run_end = char_pos + len(t)
        if first_run is None and run_end > start_char:
            first_run = i
            first_run_char_offset = start_char - char_pos
        if run_end >= end_char:
            last_run = i
            break
        char_pos = run_end

    if first_run is None or last_run is None or first_run == last_run:
        return xml  # bitta runda — yuqorida almashtirildi

    new_parts = []
    prev_end = 0
    for i, m in enumerate(runs):
        new_parts.append(xml[prev_end:m.start()])
        if i == first_run:
            before = texts[i][:first_run_char_offset]
            new_parts.append(m.group(1) + before + new_text + m.group(3))
        elif first_run < i <= last_run:
            new_parts.append(m.group(1) + "" + m.group(3))
        else:
            new_parts.append(m.group(0))
        prev_end = m.end()
    new_parts.append(xml[prev_end:])
    return "".join(new_parts)


def replace_in_docx(template_path, replacements):
    """
    Shablon .docx ichidagi matnlarni almashtiradi.
    Barcha formatlash (shrift, rang, o'lcham, jadval) to'liq saqlanadi.
    """
    with zipfile.ZipFile(template_path, "r") as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    doc_xml = files["word/document.xml"].decode("utf-8")

    for old_text, new_text in replacements.items():
        if not old_text or not new_text:
            continue
        # Avval to'g'ridan-to'g'ri almashtir (bitta run)
        doc_xml = doc_xml.replace(old_text, new_text)
        # So'ng split runs uchun
        doc_xml = _replace_split_runs(doc_xml, old_text, new_text)

    files["word/document.xml"] = doc_xml.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, files[name])
    buf.seek(0)
    return buf.read()


def generate_docx(company, product, template_type):
    templates_path = Path(TEMPLATES_FOLDER)
    template_file = templates_path / (
        "shablon_mini.docx" if template_type == "mini" else "shablon_big.docx"
    )

    if not template_file.exists():
        raise FileNotFoundError(f"Shablon topilmadi: {template_file}")

    clean_product = re.sub(r'^[\U00010000-\U0010ffff\u2600-\u26FF\u2700-\u27BF\s]+', '', product).strip()

    logger.info(f"Generatsiya: company={company}, product={clean_product}, type={template_type}")

    replacements = {
        "Чеснок свежий": clean_product,
        "ООО «BIG CONSALT»": company["exporter_name"],
        "100000 Республика Узбекистан, г. Ташкент, Шайхантохурский район, ул. Навои, дом 37.": company["exporter_addr"],
        "ООО «Сармант-ЮГ»": company["importer_name"],
        "196006, Россия, Санкт-Петербург г., муниципальный округ Московская застава вн. тер. г., Новорощинская ул., д. 4": company["importer_addr"],
    }

    return replace_in_docx(template_file, replacements)


# ── /start ───────────────────────────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await show_company_list(update.message, context)


async def show_company_list(message, context: ContextTypes.DEFAULT_TYPE) -> None:
    companies = load_companies()
    keyboard = []
    for c in companies:
        keyboard.append([
            InlineKeyboardButton(c["exporter_name"], callback_data=f"company:{c['id']}"),
            InlineKeyboardButton("🗑", callback_data=f"delete:{c['id']}"),
        ])
    keyboard.append([InlineKeyboardButton("➕ Firma qo'shish", callback_data="add_company")])
    await message.reply_text(
        "🏢 Qaysi firma uchun markitovka kerak?\n\nFirmani tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


# ── Firma tanlandi ────────────────────────────────────────────────────────────

async def company_selected(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    company_id = query.data.split(":", 1)[1]
    company = get_company_by_id(company_id)
    if not company:
        await query.edit_message_text("Firma topilmadi. /start bosing.")
        return
    context.user_data["company_id"] = company_id
    keyboard = [
        [InlineKeyboardButton("🖨 Mini printer (termal)", callback_data="printer:mini")],
        [InlineKeyboardButton("🖨 Katta printer (A4)", callback_data="printer:big")],
        [InlineKeyboardButton("⬅️ Orqaga", callback_data="back:start")],
    ]
    await query.edit_message_text(
        f"🏢 <b>{company['exporter_name']}</b>\n\nQaysi printer uchun?",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="HTML",
    )


# ── Printer tanlandi ──────────────────────────────────────────────────────────

async def printer_selected(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    printer_type = query.data.split(":", 1)[1]
    context.user_data["printer_type"] = printer_type

    keyboard = []
    row = []
    for i, product in enumerate(PRODUCTS):
        row.append(InlineKeyboardButton(product, callback_data=f"product:{i}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)

    label = "Mini printer" if printer_type == "mini" else "Katta printer"
    await query.edit_message_text(
        f"🖨 <b>{label}</b>\n\nQaysi tovar uchun markitovka?",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="HTML",
    )


# ── Tovar tanlandi ────────────────────────────────────────────────────────────

async def product_selected(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    product_index = int(query.data.split(":", 1)[1])
    product = PRODUCTS[product_index]
    company_id = context.user_data.get("company_id")
    printer_type = context.user_data.get("printer_type", "big")
    company = get_company_by_id(company_id)
    if not company:
        await query.edit_message_text("Firma topilmadi. /start bosing.")
        return

    await query.edit_message_text(
        f"⏳ <b>{product}</b> markitovkasi tayyorlanmoqda...",
        parse_mode="HTML",
    )

    try:
        docx_bytes = generate_docx(company, product, printer_type)
    except FileNotFoundError as e:
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"⚠️ {e}",
            parse_mode="HTML",
        )
        return
    except Exception as e:
        logger.exception("Docx yaratishda xatolik")
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"⚠️ Xatolik: {e}",
        )
        return

    printer_label = "mini" if printer_type == "mini" else "katta"
    filename = f"Markitovka_{product}_{printer_label}.docx"
    await context.bot.send_document(
        chat_id=query.message.chat_id,
        document=io.BytesIO(docx_bytes),
        filename=filename,
        caption=(
            f"✅ <b>{company['exporter_name']}</b>\n"
            f"📦 Tovar: <b>{product}</b>\n"
            f"🖨 Printer: <b>{printer_label}</b>"
        ),
        parse_mode="HTML",
    )

    keyboard = [[InlineKeyboardButton("🔄 Yana boshdan", callback_data="back:start")]]
    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text="Boshqa markitovka kerakmi?",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    logger.info(f"Yuborildi: {company['exporter_name']} / {product} / {printer_type}")


# ── Firma qo'shish ────────────────────────────────────────────────────────────


async def add_company_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    
    logger.info(f"Firma qo'shish bosdi: {query.from_user.id}, ADMIN_ID: {ADMIN_ID}")
    
    if query.from_user.id not in ADMIN_ID:
        await query.edit_message_text(
            "⛔ Siz admin emassiz. Bu amal faqat admin uchun.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("⬅️ Orqaga", callback_data="back:start")
            ]])
        )
        return ConversationHandler.END
    
    await query.edit_message_text(
        "🏢 <b>Yangi firma qo'shish</b>\n\n"
        "1️⃣ <b>Exporter</b> (Ishlab chiqaruvchi) firma nomini kiriting:\n\n"
        "Bekor qilish: /bekor",
        parse_mode="HTML",
    )
    return ADD_EXPORTER_NAME


async def add_exporter_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["new_exporter_name"] = update.message.text.strip()
    await update.message.reply_text(
        "📍 Exporter <b>manzilini</b> kiriting\n(shahar, ko'cha, uy raqami):",
        parse_mode="HTML",
    )
    return ADD_EXPORTER_ADDR


async def add_exporter_addr(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["new_exporter_addr"] = update.message.text.strip()
    await update.message.reply_text(
        "2️⃣ <b>Importer</b> (Xaridor) firma nomini kiriting:",
        parse_mode="HTML",
    )
    return ADD_IMPORTER_NAME


async def add_importer_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["new_importer_name"] = update.message.text.strip()
    await update.message.reply_text(
        "📍 Importer <b>manzilini</b> kiriting:",
        parse_mode="HTML",
    )
    return ADD_IMPORTER_ADDR


async def add_importer_addr(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    importer_addr = update.message.text.strip()
    companies = load_companies()
    new_company = {
        "id": f"c{len(companies)+1}_{abs(hash(context.user_data['new_exporter_name'])) % 100000}",
        "exporter_name": context.user_data["new_exporter_name"],
        "exporter_addr": context.user_data["new_exporter_addr"],
        "importer_name": context.user_data["new_importer_name"],
        "importer_addr": importer_addr,
    }
    companies.append(new_company)
    save_companies(companies)

    await update.message.reply_text(
        f"✅ <b>{new_company['exporter_name']}</b> qo'shildi!\n\n"
        f"📤 Exporter: {new_company['exporter_name']}\n"
        f"📍 {new_company['exporter_addr']}\n\n"
        f"📥 Importer: {new_company['importer_name']}\n"
        f"📍 {new_company['importer_addr']}",
        parse_mode="HTML",
    )
    await show_company_list(update.message, context)
    context.user_data.clear()
    return ConversationHandler.END


async def cancel_add(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("❌ Bekor qilindi.")
    await show_company_list(update.message, context)
    context.user_data.clear()
    return ConversationHandler.END


# ── Orqaga ────────────────────────────────────────────────────────────────────

async def delete_company(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()

    if query.from_user.id not in ADMIN_ID:
        await query.edit_message_text(
            "⛔ Siz admin emassiz.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("⬅️ Orqaga", callback_data="back:start")
            ]])
        )
        return

    company_id = query.data.split(":", 1)[1]
    companies = load_companies()
    company = next((c for c in companies if c["id"] == company_id), None)

    if not company:
        await query.edit_message_text("⚠️ Firma topilmadi.")
        return

    # Tasdiqlash tugmasi
    keyboard = [
        [InlineKeyboardButton("✅ Ha, o'chir", callback_data=f"confirm_delete:{company_id}")],
        [InlineKeyboardButton("❌ Bekor qilish", callback_data="back:start")],
    ]
    await query.edit_message_text(
        f"⚠️ <b>{company['exporter_name']}</b> ni o'chirishni tasdiqlaysizmi?",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="HTML",
    )


async def confirm_delete_company(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()

    company_id = query.data.split(":", 1)[1]
    companies = load_companies()
    company = next((c for c in companies if c["id"] == company_id), None)
    companies = [c for c in companies if c["id"] != company_id]
    save_companies(companies)

    name = company["exporter_name"] if company else "Firma"
    await query.edit_message_text(f"✅ <b>{name}</b> o'chirildi.", parse_mode="HTML")
    await show_company_list(query.message, context)

async def back_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    target = query.data.split(":", 1)[1]

    if target == "start":
        companies = load_companies()
        keyboard = []
        for c in companies:
            keyboard.append([InlineKeyboardButton(
                c["exporter_name"], callback_data=f"company:{c['id']}"
            )])
        keyboard.append([InlineKeyboardButton("➕ Firma qo'shish", callback_data="add_company")])
        await query.edit_message_text(
            "🏢 Qaysi firma uchun markitovka kerak?\n\nFirmani tanlang:",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
    elif target == "printer":
        company_id = context.user_data.get("company_id")
        company = get_company_by_id(company_id)
        name = company["exporter_name"] if company else "Firma"
        keyboard = [
            [InlineKeyboardButton("🖨 Mini printer (termal)", callback_data="printer:mini")],
            [InlineKeyboardButton("🖨 Katta printer (A4)", callback_data="printer:big")],
            [InlineKeyboardButton("⬅️ Orqaga", callback_data="back:start")],
        ]
        await query.edit_message_text(
            f"🏢 <b>{name}</b>\n\nQaysi printer uchun?",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="HTML",
        )


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error("Xatolik:", exc_info=context.error)


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    app = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CallbackQueryHandler(add_company_start, pattern="^add_company$")],
        states={
            ADD_EXPORTER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_exporter_name)],
            ADD_EXPORTER_ADDR: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_exporter_addr)],
            ADD_IMPORTER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_importer_name)],
            ADD_IMPORTER_ADDR: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_importer_addr)],
        },
        fallbacks=[CommandHandler("bekor", cancel_add)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)
    app.add_handler(CallbackQueryHandler(company_selected, pattern=r"^company:"))
    app.add_handler(CallbackQueryHandler(printer_selected, pattern=r"^printer:"))
    app.add_handler(CallbackQueryHandler(product_selected, pattern=r"^product:"))
    app.add_handler(CallbackQueryHandler(delete_company, pattern=r"^delete:"))
    app.add_handler(CallbackQueryHandler(confirm_delete_company, pattern=r"^confirm_delete:"))
    app.add_handler(CallbackQueryHandler(back_handler, pattern=r"^back:"))
    app.add_error_handler(error_handler)

    logger.info("Bot ishga tushdi...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
