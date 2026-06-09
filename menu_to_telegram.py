import os
import datetime as dt
import requests
from openpyxl import load_workbook

# === AYARLAR ===
CHAT_ID = int(os.getenv("TELEGRAM_CHAT_ID", "1677402217"))  
GROUP_CHAT_ID = int(os.getenv("GROUP_CHAT_ID", "0"))        

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
EXCEL_PATH = os.getenv("EXCEL_PATH", "BASAK_Yemek_Listesi.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "")
SEND_TOMORROW = os.getenv("SEND_TOMORROW", "false").lower() == "true"
# =================

def parse_tr_date(s):
    if not s:
        return None
        
    # Eğer Excel zaten bir tarih objesi döndürdüyse
    if isinstance(s, (dt.datetime, dt.date)):
        return s.date() if isinstance(s, dt.datetime) else s

    # Excel'deki noktalı formatı (Örn: "9.06.2026") ayrıştır
    parts = str(s).strip().split('.')
    if len(parts) == 3:
        try:
            return dt.date(int(parts[2]), int(parts[1]), int(parts[0]))
        except ValueError:
            pass # Eğer dönüştürülemezse aşağı devam et, None döner
            
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

        # Excel sütunlarına göre ilk 5 veriyi al (A:Tarih, B:Çorba, C:Ana, D:Yardımcı1, E:Yardımcı2)
        veri = list(row)[:5]
        veri += [None] * (5 - len(veri))

        t, corba, ana, yard1, yard2 = veri

        d = parse_tr_date(t)
        
        if d == target_date:
            label = "Yarın" if SEND_TOMORROW else "Bugün"
            
            # Gün ismini datetime objesinden Türkçe olarak alıyoruz
            gunler = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
            gun_ismi = gunler[d.weekday()]

            # Boş değerleri ("None") engellemek için string formatlama
            corba_str = str(corba) if corba else "Yok"
            ana_str = str(ana) if ana else "Yok"
            yard1_str = str(yard1) if yard1 else "Yok"
            yard2_str = str(yard2) if yard2 else "Yok"

            return (
                f"🍽️ BAŞAK TRAKTÖR\n"
                f"📌 {d.strftime('%d.%m.%Y')} {gun_ismi} - {label} Yemek Menüsü\n"
                f"🍲 Çorba: {corba_str}\n"
                f"🍝 Ana Yemek: {ana_str}\n"
                f"🍖 Yardımcı: {yard1_str}\n"
                f"🥗 Salata/Tatlı: {yard2_str}"
            )

    return f"❗ {target_date.strftime('%d.%m.%Y')} tarihi için yemek menüsü bulunamadı."

def send_telegram(msg, chat_id):
    if not TOKEN:
        print("HATA: TELEGRAM_BOT_TOKEN bulunamadı.")
        return
        
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    try:
        r = requests.post(url, json={"chat_id": chat_id, "text": msg}, timeout=20)
        r.raise_for_status()
        print(f"Mesaj başarıyla gönderildi (ID: {chat_id})")
    except Exception as e:
        print(f"Mesaj gönderilirken hata oluştu (ID: {chat_id}): {e}")

if __name__ == "__main__":
    turkey_tz = dt.timezone(dt.timedelta(hours=3))
    today = dt.datetime.now(turkey_tz).date()
    
    target = today + dt.timedelta(days=1) if SEND_TOMORROW else today

    if target.weekday() >= 5:
        print("Hafta sonu - mesaj gönderilmedi.")
        raise SystemExit(0)

    msg = get_menu(target)
    
    send_telegram(msg, CHAT_ID)
    
    if not msg.startswith("❗") and GROUP_CHAT_ID != 0:
        send_telegram(msg, GROUP_CHAT_ID)
    else:
        print("Hata, tatil günü veya grup ID'si 0 olduğu için gruba mesaj atılmadı.")
