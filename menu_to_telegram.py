import os
import datetime as dt
import requests
from openpyxl import load_workbook

# === AYARLAR ===
CHAT_ID = int(os.getenv("TELEGRAM_CHAT_ID", "1677402217"))  # Sana gelecek özel mesaj
GROUP_CHAT_ID = int(os.getenv("GROUP_CHAT_ID", "0"))        # BAŞAK YEMEK MENÜ grubuna gidecek mesaj

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
EXCEL_PATH = os.getenv("EXCEL_PATH", "BASAK_Yemek_Listesi.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "")
SEND_TOMORROW = os.getenv("SEND_TOMORROW", "false").lower() == "true"
# =================

MONTHS = {
    "Ocak": 1, "Şubat": 2, "Mart": 3, "Nisan": 4, "Mayıs": 5, "Haziran": 6,
    "Temmuz": 7, "Ağustos": 8, "Eylül": 9, "Ekim": 10, "Kasım": 11, "Aralık": 12
}

def parse_tr_date(s):
    if not s:
        return None
    parts = str(s).strip().split()
    if len(parts) != 3:
        return None
    try:
        return dt.date(int(parts[2]), MONTHS.get(parts[1]), int(parts[0]))
    except:
        return None

def get_menu(target_date):
    try:
        wb = load_workbook(EXCEL_PATH, data_only=True)
        if SHEET_NAME:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.active
    except Exception as e:
        return f"❗ Excel dosyası veya sayfası bulunamadı: {e}"

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        veri = list(row)[:6]
        veri += [None] * (6 - len(veri))
        t, gun, corba, ana, yard, tatli = veri
        d = parse_tr_date(t)
        if d == target_date:
            label = "Yarın" if SEND_TOMORROW else "Bugün"
            return (
                f"🍽️ BAŞAK TRAKTÖR\n"
                f"📌 {gun} {label} Yemek Menüsü\n"
                f"🍲 Çorba: {corba}\n"
                f"🍝 Ana Yemek: {ana}\n"
                f"🍖 Yardımcı: {yard}\n"
                f"🍮 Tatlı/Meyve: {tatli}"
            )
    return f"❗ Bugün için yemek menüsü bulunamadı."


def send_telegram(chat_id, msg):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    payload = {"chat_id": chat_id, "text": msg}
    response = requests.post(url, json=payload, timeout=10)
    response.raise_for_status()


def main():
    today = dt.date.today()
    if today.weekday() >= 5:
        print("Hafta sonu - mesaj gönderilmedi.")
        return
    target_date = today
    if SEND_TOMORROW:
        target_date += dt.timedelta(days=1)
    menu = get_menu(target_date)
    send_telegram(CHAT_ID, menu)
    if not menu.startswith("❗") and GROUP_CHAT_ID != 0:
        send_telegram(GROUP_CHAT_ID, menu)


if __name__ == "__main__":
    main()
