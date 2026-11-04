# -*- coding: utf-8 -*-
import os, re, json, urllib.request
from datetime import datetime, timedelta

base_dir = os.path.dirname(os.path.abspath(__file__))
main_path = os.path.join(base_dir, "main.py")

with open(main_path, "r", encoding="utf-8") as f:
    code = f.read()

print("\n--- 1. MAIN.PY ÇNDEK KRA / TÜFE KUTULARI ---")
for m in re.finditer(r'(self\.[a-zA-Z0-9_]+)\s*=\s*QLineEdit\(([^\)]*)\)', code):
    var_name, init_val = m.group(1), m.group(2)
    if any(k in var_name.lower() for k in ["kr", "kira", "tufe", "tuik", "oran", "artis", "enflasyon"]):
        print(f"Değişken: {var_name:25s} | Başlangıç Değeri: {init_val}")

print("\n--- 2. KRA SEKMES METODU (build_cat) ÇER ---")
for m in re.finditer(r'def (build_cat[0-9a-zA-Z_]*kira[0-9a-zA-Z_]*)\(self\):([\s\S]*?)(?=\n    def |\Z)', code, re.IGNORECASE):
    print(f"Metod Adı: {m.group(1)}")
    lines = [line.strip() for line in m.group(2).splitlines() if any(x in line.lower() for x in ["qlineedit", "tufe", "tüfe", "oran", "kira", "cfg"])]
    for l in lines[:15]:
        print("  ->", l)

# EVDS API Key bul ve iki oranı da canlı hesapla
api_key = ""
for fn in ["mevzuat_ayarlari.json", "config.json", "tcmb_mcp_config.json"]:
    fp = os.path.join(base_dir, fn)
    if os.path.exists(fp):
        try:
            d = json.load(open(fp, "r", encoding="utf-8"))
            for k in ["evds_api_key", "tcmb_api_key", "api_key", "key"]:
                if d.get(k): api_key = str(d[k]).strip()
        except Exception: pass

if not api_key:
    mk = re.search(r'["\']key["\']\s*:\s*["\']([A-Za-z0-9]{8,16})["\']', code)
    if mk: api_key = mk.group(1)

print("\n--- 3. TCMB EVDS CANLI TÜFE HESAPLAMASI ---")
if api_key:
    try:
        end_d = datetime.now()
        start_d = end_d - timedelta(days=850)
        url = f"https://evds2.tcmb.gov.tr/service/evds/series=TP.FG.J0&startDate={start_d.strftime('%d-%m-%Y')}&endDate={end_d.strftime('%d-%m-%Y')}&type=json&formulas=0&frequency=5"
        req = urllib.request.Request(url, headers={"key": api_key, "User-Agent": "Mozilla/5.0"})
        data = json.loads(urllib.request.urlopen(req, timeout=5).read().decode("utf-8"))
        vals = [(it.get("Tarih"), float(str(it.get("TP_FG_J0")).replace(",", "."))) for it in data.get("items", []) if it.get("TP_FG_J0") not in (None, "", "ND")]
        if len(vals) >= 24:
            son_ay_tarih, son_ay_endeks = vals[-1]
            gecen_yil_tarih, gecen_yil_endeks = vals[-13]
            yoy_degisim = ((son_ay_endeks / gecen_yil_endeks) - 1.0) * 100.0
            oniki_ay_ort = ((sum(v[1] for v in vals[-12:]) / sum(v[1] for v in vals[-24:-12])) - 1.0) * 100.0
            print(f"Son Açıklanan Ay ({son_ay_tarih}) Endeksi : {son_ay_endeks}")
            print(f"Geçen Yıl Aynı Ay ({gecen_yil_tarih}) Endeksi: {gecen_yil_endeks}")
            print(f"A) Bir Önceki Yılın Aynı Ayına Göre (12 Aylık Yıllık Değişim): %{yoy_degisim:.2f}")
            print(f"B) On ki Aylık Ortalamalara Göre Değişim (TBK m.344 Kira)  : %{oniki_ay_ort:.2f}")
    except Exception as e:
        print("EVDS Bağlantı Notu:", e)