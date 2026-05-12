# Markitovka Bot — O'rnatish qo'llanmasi

## Papka tuzilishi

```
markitovka-bot/
├── bot.py
├── config.py
├── requirements.txt
├── firmalar.json        ← bot o'zi yaratadi
└── shablonlar/
    ├── shablon_mini.docx   ← sizning mini printer shabloni
    └── shablon_big.docx    ← sizning katta printer shabloni
```

---

## 1-qadam — Kutubxona o'rnatish

```
pip install -r requirements.txt
```

---

## 2-qadam — Shablonlarni joylashtirish

`shablonlar/` papkasiga ikki fayl qo'ying:
- `shablon_mini.docx` — termal/mini printer uchun
- `shablon_big.docx` — katta (A4) printer uchun

Shablonlarda quyidagi matnlar **avtomatik almashtiriladi**:
| Shablondagi matn | Nima bo'ladi |
|---|---|
| `Чеснок свежий` | Tanlangan tovar nomi |
| `ООО «BIG CONSALT»` | Exporter firma nomi |
| Shayxontohur, Navoi 37 | Exporter manzili |
| `ООО «Сармант-ЮГ»` | Importer firma nomi |
| Sankt-Peterburg adres | Importer manzili |

---

## 3-qadam — Token kiriting

`config.py` faylini oching:
```python
BOT_TOKEN = "sizning_tokeningiz_bu_yerga"
```

---

## 4-qadam — Botni ishga tushirish

```
python bot.py
```

---

## Bot qanday ishlaydi

```
/start
  └─► Firmalar ro'yxati  +  [➕ Firma qo'shish]
        └─► Firma tanlanganda:
              └─► Printer tanlash: Mini / Katta
                    └─► Tovar tanlash (22 ta)
                          └─► Shablon ichidagi matn almashtiriladi
                                └─► .docx fayl yuboriladi
```

## Firma qo'shish

Bot ichidan "➕ Firma qo'shish" tugmasini bosing:
1. Exporter firma nomi
2. Exporter manzili
3. Importer firma nomi
4. Importer manzili

Firmalar `firmalar.json` faylida saqlanadi.

## Tovarlar ro'yxatini o'zgartirish

`bot.py` faylida `PRODUCTS` listini tahrirlang:
```python
PRODUCTS = [
    "Olma", "Nok", "Gilos", ...
]
```
