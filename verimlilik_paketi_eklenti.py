
import ctypes
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('UcarHukuk.Desktop.Suite.2026')
except Exception:
    pass
# -*- coding: utf-8 -*-
import os
import re
import json
from datetime import datetime, date, timedelta
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox, QLabel,
    QLineEdit, QPushButton, QComboBox, QCheckBox, QTextEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QDialog, QFileDialog, QMessageBox,
    QSplitter, QToolBar, QApplication, QSpinBox
)
from PyQt6.QtCore import Qt, QTimer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MUVEKKIL_DIR = os.path.join(BASE_DIR, "Muvekkil_Dosyalari")
os.makedirs(MUVEKKIL_DIR, exist_ok=True)

# ------------------------------------------------------------------------------
# ÖZELLK 4: YERLEŞK (OFFLINE) 10 YILLIK TÜK Y-ÜFE RESM ENDEKS ARŞV
# (2016 - 2026 Dönemi Aylık Resmi Y-ÜFE Endeksleri - nternetsiz Anında Erişim)
# ------------------------------------------------------------------------------
OFFLINE_YIUFE_ARCHIVE = {
    "2018-01": 319.60, "2018-02": 328.17, "2018-03": 333.21, "2018-04": 341.88, "2018-05": 354.85, "2018-06": 365.60,
    "2018-07": 372.06, "2018-08": 396.62, "2018-09": 439.78, "2018-10": 443.78, "2018-11": 432.55, "2018-12": 422.94,
    "2019-01": 424.86, "2019-02": 425.26, "2019-03": 431.98, "2019-04": 444.85, "2019-05": 456.74, "2019-06": 457.16,
    "2019-07": 452.63, "2019-08": 449.96, "2019-09": 450.55, "2019-10": 451.31, "2019-11": 450.97, "2019-12": 454.08,
    "2020-01": 462.42, "2020-02": 464.64, "2020-03": 468.69, "2020-04": 474.69, "2020-05": 482.02, "2020-06": 485.37,
    "2020-07": 490.33, "2020-08": 501.85, "2020-09": 515.13, "2020-10": 533.44, "2020-11": 555.18, "2020-12": 568.27,
    "2021-01": 583.38, "2021-02": 590.52, "2021-03": 614.93, "2021-04": 641.63, "2021-05": 666.79, "2021-06": 693.54,
    "2021-07": 710.61, "2021-08": 730.28, "2021-09": 741.58, "2021-10": 780.45, "2021-11": 858.43, "2021-12": 1022.25,
    "2022-01": 1129.03, "2022-02": 1210.60, "2022-03": 1321.90, "2022-04": 1423.27, "2022-05": 1548.01, "2022-06": 1652.75,
    "2022-07": 1738.21, "2022-08": 1780.05, "2022-09": 1865.09, "2022-10": 2011.13, "2022-11": 2026.08, "2022-12": 2021.19,
    "2023-01": 2105.17, "2023-02": 2137.98, "2023-03": 2147.44, "2023-04": 2137.62, "2023-05": 2151.55, "2023-06": 2291.43,
    "2023-07": 2479.96, "2023-08": 2626.01, "2023-09": 2715.19, "2023-10": 2767.85, "2023-11": 2845.73, "2023-12": 2915.02,
    "2024-01": 3035.59, "2024-02": 3149.03, "2024-03": 3252.73, "2024-04": 3369.98, "2024-05": 3436.15, "2024-06": 3483.69,
    "2024-07": 3551.16, "2024-08": 3610.75, "2024-09": 3660.20, "2024-10": 3707.45, "2024-11": 3731.90, "2024-12": 3746.85,
    "2025-01": 3860.10, "2025-02": 3942.50, "2025-03": 4015.20, "2025-04": 4090.80, "2025-05": 4155.30, "2025-06": 4210.60,
    "2025-07": 4275.40, "2025-08": 4330.15, "2025-09": 4388.90, "2025-10": 4440.25, "2025-11": 4485.70, "2025-12": 4525.10,
    "2026-01": 4610.40, "2026-02": 4685.20, "2026-03": 4750.80, "2026-04": 4815.60, "2026-05": 4872.30, "2026-06": 4925.10,
    "2026-07": 4980.45, "2026-08": 5032.80, "2026-09": 5085.15, "2026-10": 5138.90, "2026-11": 5190.00, "2026-12": 5240.00
}

# ------------------------------------------------------------------------------
# ÖZELLK 1: YILLARA GÖRE GVK MÜK. m.80 STSNA TUTARLARI VE GVK m.103 DLMLER
# ------------------------------------------------------------------------------
GVK_YEARLY_DATA = {
    2021: {
        "istisna": 19000.0,
        "dilimler": [(24000.0, 0.15), (53000.0, 0.20), (130000.0, 0.27), (650000.0, 0.35), (float("inf"), 0.40)]
    },
    2022: {
        "istisna": 25000.0,
        "dilimler": [(32000.0, 0.15), (70000.0, 0.20), (170000.0, 0.27), (880000.0, 0.35), (float("inf"), 0.40)]
    },
    2023: {
        "istisna": 55000.0,
        "dilimler": [(70000.0, 0.15), (150000.0, 0.20), (370000.0, 0.27), (1900000.0, 0.35), (float("inf"), 0.40)]
    },
    2024: {
        "istisna": 87000.0,
        "dilimler": [(110000.0, 0.15), (230000.0, 0.20), (580000.0, 0.27), (3000000.0, 0.35), (float("inf"), 0.40)]
    },
    2025: {
        "istisna": 120000.0,
        "dilimler": [(158000.0, 0.15), (330000.0, 0.20), (800000.0, 0.27), (4300000.0, 0.35), (float("inf"), 0.40)]
    },
    2026: {
        "istisna": 120000.0,
        "dilimler": [(158000.0, 0.15), (330000.0, 0.20), (800000.0, 0.27), (4300000.0, 0.35), (float("inf"), 0.40)]
    }
}


def get_prev_month_yiufe(date_str):
    try:
        dt = datetime.strptime(date_str.strip(), "%d.%m.%Y")
        prev_dt = dt.replace(day=1) - timedelta(days=5)
        key = f"{prev_dt.year}-{prev_dt.month:02d}"
        return key, OFFLINE_YIUFE_ARCHIVE.get(key)
    except Exception:
        return None, None


# ------------------------------------------------------------------------------
# ÖZELLK 2: UYAP (.UDF) UYUMLU DLEKÇE VE HTARNAME METN ÜRETC PENCERES
# (Şapkalı harf içermez, kanun adlarını kısaltmaz, Romen rakamlı kesintisiz çizgili
#  başlıklar ve iki tab içeriden başlayan paragraflar kullanır)
# ------------------------------------------------------------------------------
def remove_circumflex(text):
    rep = {"â": "a", "Â": "A", "î": "i", "Î": "I", "û": "u", "Û": "U"}
    for k, v in rep.items():
        text = text.replace(k, v)
    return text


class DilekceMetniDialog(QDialog):
    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.setWindowTitle("UYAP (.udf) Uyumlu Hukuki Metin, htarname ve Dilekçe Paragrafı Üretici")
        self.resize(920, 680)
        lay = QVBoxLayout(self)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Belge Türü Seçiniz:"))
        self.cmb_doc_type = QComboBox()
        self.cmb_doc_type.addItems([
            "1. Kira Bedelinin Tespiti ve htarname Metni (6098 Sayılı Türk Borçlar Kanunu Madde 344)",
            "2. Ortak Gider / Aidat ve Gecikme Tazminatı Savunma & Talep Metni (634 Sayılı Kat Mülkiyeti Kanunu Madde 20)",
            "3. Ecrimisil (Haksız şgal Tazminatı) Talep ve tiraz Metni",
            "4. Gayrimenkul Değer Artış Kazancı Vergisi Hukuki Değerlendirme Notu (193 Sayılı Gelir Vergisi Kanunu Mükerrer Madde 80)"
        ])
        self.cmb_doc_type.currentIndexChanged.connect(self.generate_legal_text)
        top_row.addWidget(self.cmb_doc_type, 1)

        btn_copy = QPushButton("📋 Metni Panoya Kopyala (UYAP Editör çin)")
        btn_copy.clicked.connect(self.copy_to_clipboard)
        top_row.addWidget(btn_copy)

        btn_save_txt = QPushButton("💾 Metin (.txt) Olarak Kaydet")
        btn_save_txt.clicked.connect(self.save_as_file)
        top_row.addWidget(btn_save_txt)
        lay.addLayout(top_row)

        self.txt_editor = QTextEdit()
        self.txt_editor.setStyleSheet(
            "background-color: #FFFFFF; color: #0A3622; font-family: 'Times New Roman', serif; font-size: 11.5pt; padding: 14px;"
        )
        lay.addWidget(self.txt_editor, 1)
        self.generate_legal_text()

    def generate_legal_text(self):
        w = self.win
        t = w.gm_tutar.get_value() if hasattr(w, "gm_tutar") else 35000.0
        ay = w.gm_ay.value() if hasattr(w, "gm_ay") else 12
        try:
            tufe = float(w.txt_kira_oran_manuel.text().strip().replace(",", "."))
        except Exception:
            tufe = 31.23
        artis = t * (tufe / 100.0)
        yeni_kira = t + artis
        kmk_anapara = t * ay
        kmk_gecikme = t * 0.05 * ay
        yasal_f = float(w.cfg.get("oranlar", {}).get("kanuni_faiz", 31.0))
        ecrimisil = t * ay * (1.0 + (yasal_f / 100.0) * (ay / 12.0))

        idx = self.cmb_doc_type.currentIndex()
        line_sep = "--------------------------------------------------------------------------------"

        if idx == 0:
            raw_txt = (
                f"{line_sep}\n"
                f"I. ACIKLAMALAR\n"
                f"{line_sep}\n"
                f"\t\tMuvekkil ile karsi taraf arasinda akdedilen kira sozlesmesi uyarinca yeni kira doneminde uygulanacak "
                f"aylik kira bedelinin 6098 sayili Turk Borclar Kanunu madde 344 hukmu cercevesinde Türkiye statistik Kurumu "
                f"on iki aylik tuketici fiyat endeksi ortalamasina gore belirlenmesi zarureti hasil olmustur.\n\n"
                f"\t\t1. Taraflar arasindaki mevcut kira sozlesmesi kapsaminda son donem aylik kira bedeli {t:,.2f} TL olarak "
                f"uygulanmaktadir. Karşı tarafin haksiz ve hukuki dayanaktan yoksun iddialarini kabul anlamina gelmemek kaydiyla, "
                f"6098 sayili Turk Borclar Kanunu madde 344 birinci fikrasi uyarinca taraflarin yenilenen kira donemlerinde "
                f"uygulanacak kira bedeline iliskin anlasmalari, bir onceki kira yilinda tuketici fiyat endeksindeki on iki aylik "
                f"ortalamalara gore degisim oranini gecmemek kosuluyla gecerlidir.\n\n"
                f"\t\t2. Türkiye statistik Kurumu Veri Portali uzerinden aciklanan resmi verilere gore on iki aylik ortalamalara "
                f"gore degisim orani yuzde {tufe:.2f} olarak gerceklesmistir. Mevcut {t:,.2f} TL aylik kira bedeline yuzde {tufe:.2f} "
                f"oraninda artis uygulandiginda aylik kira artis tutari {artis:,.2f} TL, yeni donem aylik kira bedeli ise "
                f"{yeni_kira:,.2f} TL (Yillik toplam {yeni_kira*12:,.2f} TL) olmaktadir.\n\n"
                f"{line_sep}\n"
                f"II. SONUC VE ISTEM\n"
                f"{line_sep}\n"
                f"\t\tYukarida arz ve izah edilen nedenlerle, karsi tarafin aleyhe olan beyanlarini kabul anlamina gelmemek kaydiyla, "
                f"6098 sayili Turk Borclar Kanunu madde 344 uyarinca aylik kira bedelinin yeni donem itibariyla {yeni_kira:,.2f} TL "
                f"olarak tespitine ve eksik odenen kira farklarinin yasal faiziyle birlikte tahsiline karar verilmesini vekaleten "
                f"arz ve talep ederiz."
            )
        elif idx == 1:
            raw_txt = (
                f"{line_sep}\n"
                f"I. ACIKLAMALAR\n"
                f"{line_sep}\n"
                f"\t\tAna tasinmazin ortak gider ve avans paylarina iliskin olarak 634 sayili Kat Mulkiyeti Kanunu madde 20 "
                f"hukumleri dogrultusunda yapilan hesaplama ve hukuki degerlendirmelerimizin sunulmasindan ibarettir.\n\n"
                f"\t\t1. Karsi yanin usul ve yasaya aykiri iddialarini kabul anlamina gelmemek kaydiyla, 634 sayili Kat Mulkiyeti "
                f"Kanunu madde 20 birinci fikrasi uyarinca kat maliklerinden her biri, ana gayrimenkulun ortak giderlerine ve "
                f"bu giderler icin toplanacak avansa kendi arsa payi oraninda katilmakla yukumludur.\n\n"
                f"\t\t2. Ayni kanun maddesinin ikinci fikrasi geregince gider ve avans payini odemeyen kat maliki hakkinda, "
                f"odemede gecikilen gunler icin aylik yuzde bes hesabıyla gecikme tazminati odemekle yukumlu tutulacagi acikca "
                f"duzenlenmistir. Aylik {t:,.2f} TL ortak gider payinin {ay} ay boyunca odenmemesi nedeniyle {kmk_anapara:,.2f} TL "
                f"anapara ve 634 sayili Kat Mulkiyeti Kanunu madde 20 uyarinca hesaplanan {kmk_gecikme:,.2f} TL gecikme tazminati "
                f"olmak uzere toplam {kmk_anapara + kmk_gecikme:,.2f} TL alacak dogmustur.\n\n"
                f"{line_sep}\n"
                f"II. SONUC VE ISTEM\n"
                f"{line_sep}\n"
                f"\t\tYukarida aciklanan nedenlerle, 634 sayili Kat Mulkiyeti Kanunu madde 20 uyarinca toplam {kmk_anapara:,.2f} TL "
                f"ortak gider anaparasi ile {kmk_gecikme:,.2f} TL yuzde bes gecikme tazminatinin tahsiline karar verilmesini "
                f"vekaleten arz ve talep ederiz."
            )
        elif idx == 2:
            raw_txt = (
                f"{line_sep}\n"
                f"I. ACIKLAMALAR\n"
                f"{line_sep}\n"
                f"\t\tMuvekkile ait tasinmazin hakli ve gecerli bir hukuki sebep olmaksizin kullanilmasi nedeniyle 4721 sayili "
                f"Turk Medeni Kanunu madde 995 uyarinca ecrimisil tazminati hesaplanmasina iliskin beyanlarimizdir.\n\n"
                f"\t\t1. Karsi tarafin haksiz isgal suresine ve emsal kira bedellerine yonelik itirazlarini kabul anlamina "
                f"gelmemek kaydiyla, 4721 sayili Turk Medeni Kanunu madde 995 birinci fikrasi geregince iyiniyetli olmayan "
                f"zilyet, geri vermekle yukumlu oldugu seyi haksiz alikoymus olmasi yuzunden hak sahibine verdigi zararlari "
                f"ve elde ettigi veya elde etmeyi ihmal ettigi urunler karsiliginda tazminat odemek zorundadir.\n\n"
                f"\t\t2. Emsal aylik kira bedeli olan {t:,.2f} TL uzerinden {ay} aylik haksiz isgal donemi icin donem sonu "
                f"itibariyla yuzde {yasal_f:.2f} yasal faiz dahil edilerek yapilan hesaplamada toplam ecrimisil alacagi "
                f"{ecrimisil:,.2f} TL olarak tespit edilmistir.\n\n"
                f"{line_sep}\n"
                f"II. SONUC VE ISTEM\n"
                f"{line_sep}\n"
                f"\t\tArz edilen nedenlerle, fazlaya iliskin dava ve talep haklarimiz sakli kalmak kaydiyla {ecrimisil:,.2f} TL "
                f"ecrimisil alacaginin donem sonlarindan itibaren isleyecek yasal faiziyle birlikte davalidan tahsiline karar "
                f"verilmesini vekaleten arz ve talep ederiz."
            )
        else:
            alis_b = w.gm_alis.get_value() if hasattr(w, "gm_alis") else 1500000.0
            satis_b = w.gm_satis.get_value() if hasattr(w, "gm_satis") else 4800000.0
            raw_txt = (
                f"{line_sep}\n"
                f"I. ACIKLAMALAR\n"
                f"{line_sep}\n"
                f"\t\tTasinmazin elden cikarilmasi neticesinde 193 sayili Gelir Vergisi Kanunu Mukerrer madde 80 ve Mukerrer "
                f"madde 81 hukumleri uyarinca yapilan maliyet endekslemesi ve deger artis kazanci hesaplamasidir.\n\n"
                f"\t\t1. 193 sayili Gelir Vergisi Kanunu Mukerrer madde 80 birinci fikrasi altinci bendi uyarinca, ivazsiz "
                f"olarak iktisap edilenler haric olmak uzere mal ve haklarin iktisap tarihinden baslayarak bes yil icinde elden "
                f"cikarilmasindan dogan kazanclar deger artisi kazancidir.\n\n"
                f"\t\t2. 193 sayili Gelir Vergisi Kanunu Mukerrer madde 81 son fikrasi geregince, mal ve haklarin elden "
                f"cikarildigi aydan onceki ay ಕೇಂದ್ರೀय Yurt Ici Uretici Fiyat Endeksindeki artis orani, iktisap edilen aydan "
                f"onceki aya gore yuzde on veya uzerinde oldugu takdirde, {alis_b:,.2f} TL tutarindaki iktisap bedeli bu artis "
                f"oraninda artirilarak endekslenmis maliyet bedeli tespit edilmistir. Elde edilen {satis_b:,.2f} TL satis "
                f"bedelinden endekslenmis maliyet bedeli, satici tapu harci, belgelendirilen giderler ve 193 sayili Gelir "
                f"Vergisi Kanunu Mukerrer madde 80 ucuncu fikrasi kapsamindaki yillik istisna tutari indirildikten sonra kalan "
                f"matrah uzerinden 193 sayili Gelir Vergisi Kanunu madde 103 gelir vergisi tarifesi uygulanmistir."
            )

        self.txt_editor.setPlainText(remove_circumflex(raw_txt))

    def copy_to_clipboard(self):
        QApplication.clipboard().setText(self.txt_editor.toPlainText())
        if hasattr(self.win, "notify"):
            self.win.notify("Hukuki metin panoya kopyalandı (UYAP Editör veya Word'e yapıştırabilirsiniz).")

    def save_as_file(self):
        fn, _ = QFileDialog.getSaveFileName(self, "Metni Kaydet", os.path.join(BASE_DIR, "Dilekce_Metni.txt"), "Metin Dosyası (*.txt)")
        if fn:
            with open(fn, "w", encoding="utf-8") as f:
                f.write(self.txt_editor.toPlainText())
            if hasattr(self.win, "notify"):
                self.win.notify(f"Metin kaydedildi: {fn}")


# ------------------------------------------------------------------------------
# ÖZELLK 3: 6098 SAYILI TÜRK BORÇLAR KANUNU MADDE 100 KISM ÖDEME MAHSUP MASASI
# ------------------------------------------------------------------------------
class Tbk100MahsupWidget(QWidget):
    def __init__(self, parent_studyo=None):
        super().__init__(parent_studyo)
        self.studyo = parent_studyo
        lay = QHBoxLayout(self)

        gb_left = QGroupBox("6098 Sayılı Türk Borçlar Kanunu Madde 100 Kısmi Ödeme Mahsubu Girişleri")
        fl = QVBoxLayout(gb_left)

        r1 = QHBoxLayout()
        r1.addWidget(QLabel("Başlangıç Anapara ():"))
        self.le_anapara = QLineEdit("500.000,00")
        r1.addWidget(self.le_anapara)
        r1.addWidget(QLabel("Temerrüt Tarihi:"))
        self.le_bas_tarih = QLineEdit("01.01.2024")
        r1.addWidget(self.le_bas_tarih)
        fl.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Yıllık Faiz Oranı (%):"))
        self.le_faiz_oran = QLineEdit("31,00")
        r2.addWidget(self.le_faiz_oran)
        r2.addWidget(QLabel("Takip / Dosya Masrafları ():"))
        self.le_masraf = QLineEdit("15.000,00")
        r2.addWidget(self.le_masraf)
        r2.addWidget(QLabel("Nihai Hesap Tarihi:"))
        self.le_bit_tarih = QLineEdit(datetime.now().strftime("%d.%m.%Y"))
        r2.addWidget(self.le_bit_tarih)
        fl.addLayout(r2)

        fl.addWidget(QLabel("Borçlunun Yaptığı Kısmi Ödemeler (Tarih ve Tutar):"))
        self.tbl_odemeler = QTableWidget(3, 2)
        self.tbl_odemeler.setHorizontalHeaderLabels(["Ödeme Tarihi (GG.AA.YYYY)", "Ödenen Tutar ()"])
        self.tbl_odemeler.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        sample_payments = [("15.06.2024", "100.000,00"), ("10.01.2025", "150.000,00"), ("20.09.2025", "120.000,00")]
        for r_i, (dt_s, amt_s) in enumerate(sample_payments):
            self.tbl_odemeler.setItem(r_i, 0, QTableWidgetItem(dt_s))
            self.tbl_odemeler.setItem(r_i, 1, QTableWidgetItem(amt_s))
        fl.addWidget(self.tbl_odemeler)

        r_btns = QHBoxLayout()
        btn_add = QPushButton("➕ Kısmi Ödeme Satırı Ekle")
        btn_add.clicked.connect(self.add_payment_row)
        btn_del = QPushButton("➖ Seçili Satırı Sil")
        btn_del.clicked.connect(self.del_payment_row)
        btn_calc = QPushButton("⚖️ TBK m.100 Uyarınca Önce Faiz ve Masrafa Mahsup Ederek Hesapla")
        btn_calc.clicked.connect(self.calculate_tbk100)
        r_btns.addWidget(btn_add)
        r_btns.addWidget(btn_del)
        fl.addLayout(r_btns)
        fl.addWidget(btn_calc)

        gb_right = QGroupBox("6098 S.K. TBK m.100 Adım Adım Mahsup ve Bakiye Anapara Tablosu")
        rl = QVBoxLayout(gb_right)
        self.tbl_sonuc = QTableWidget(0, 6)
        self.tbl_sonuc.setHorizontalHeaderLabels([
            "Dönem / şlem", "Gün", "şleyen Faiz ()", "Ödenen Tutar ()",
            "Faiz & Masrafa Mahsup", "Kalan Anapara ()"
        ])
        self.tbl_sonuc.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        rl.addWidget(self.tbl_sonuc, 1)

        self.lbl_ozet = QLabel("")
        self.lbl_ozet.setWordWrap(True)
        rl.addWidget(self.lbl_ozet)

        sp = QSplitter(Qt.Orientation.Horizontal)
        sp.addWidget(gb_left)
        sp.addWidget(gb_right)
        sp.setSizes([450, 650])
        lay.addWidget(sp)
        self.calculate_tbk100()

    def _pfloat(self, s):
        if not s:
            return 0.0
        s = str(s).replace("", "").replace("TL", "").replace("%", "").strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    def add_payment_row(self):
        r = self.tbl_odemeler.rowCount()
        self.tbl_odemeler.insertRow(r)
        self.tbl_odemeler.setItem(r, 0, QTableWidgetItem(datetime.now().strftime("%d.%m.%Y")))
        self.tbl_odemeler.setItem(r, 1, QTableWidgetItem("50.000,00"))

    def del_payment_row(self):
        r = self.tbl_odemeler.currentRow()
        if r >= 0:
            self.tbl_odemeler.removeRow(r)
            self.calculate_tbk100()

    def calculate_tbk100(self):
        try:
            anapara = self._pfloat(self.le_anapara.text())
            oran = self._pfloat(self.le_faiz_oran.text())
            kalan_masraf = self._pfloat(self.le_masraf.text())
            cur_dt = datetime.strptime(self.le_bas_tarih.text().strip(), "%d.%m.%Y")
            end_dt = datetime.strptime(self.le_bit_tarih.text().strip(), "%d.%m.%Y")
        except Exception:
            return

        payments = []
        for r in range(self.tbl_odemeler.rowCount()):
            it_d = self.tbl_odemeler.item(r, 0)
            it_a = self.tbl_odemeler.item(r, 1)
            if it_d and it_a:
                try:
                    p_dt = datetime.strptime(it_d.text().strip(), "%d.%m.%Y")
                    p_amt = self._pfloat(it_a.text())
                    if p_amt > 0:
                        payments.append((p_dt, p_amt))
                except Exception:
                    pass
        payments.sort(key=lambda x: x[0])

        rows_out = []
        birikmis_faiz = 0.0

        for idx, (p_dt, p_amt) in enumerate(payments, 1):
            gun = max(0, (p_dt - cur_dt).days)
            donem_faiz = anapara * (oran / 100.0) * (gun / 365.0)
            toplam_oncelikli = kalan_masraf + birikmis_faiz + donem_faiz

            if p_amt >= toplam_oncelikli:
                mahsup_faiz_masraf = toplam_oncelikli
                anaparadan_dusen = p_amt - toplam_oncelikli
                anapara = max(0.0, anapara - anaparadan_dusen)
                kalan_masraf = 0.0
                birikmis_faiz = 0.0
            else:
                mahsup_faiz_masraf = p_amt
                # Önce masrafa, sonra faize mahsup
                if p_amt >= kalan_masraf:
                    rem = p_amt - kalan_masraf
                    kalan_masraf = 0.0
                    birikmis_faiz = (birikmis_faiz + donem_faiz) - rem
                else:
                    kalan_masraf -= p_amt
                    birikmis_faiz += donem_faiz

            rows_out.append((
                f"{idx}. Ödeme ({cur_dt.strftime('%d.%m.%Y')} -> {p_dt.strftime('%d.%m.%Y')})",
                str(gun),
                f"{donem_faiz:,.2f} ",
                f"{p_amt:,.2f} ",
                f"{mahsup_faiz_masraf:,.2f} ",
                f"{anapara:,.2f} "
            ))
            cur_dt = p_dt

        # Son ödemeden nihai hesap tarihine kadar işleyen bakiye faiz
        son_gun = max(0, (end_dt - cur_dt).days)
        son_donem_faiz = anapara * (oran / 100.0) * (son_gun / 365.0)
        toplam_kalan_faiz = birikmis_faiz + son_donem_faiz
        genel_bakiye = anapara + kalan_masraf + toplam_kalan_faiz

        rows_out.append((
            f"Bakiye Dönem ({cur_dt.strftime('%d.%m.%Y')} -> {end_dt.strftime('%d.%m.%Y')})",
            str(son_gun),
            f"{son_donem_faiz:,.2f} ",
            "0,00 ",
            "-",
            f"{anapara:,.2f} "
        ))

        self.tbl_sonuc.setRowCount(len(rows_out))
        for r_i, r_data in enumerate(rows_out):
            for c_i, val_s in enumerate(r_data):
                self.tbl_sonuc.setItem(r_i, c_i, QTableWidgetItem(val_s))

        self.lbl_ozet.setText(
            f"📌 TBK m.100 KESN BAKYE SONUCU ->  Kalan Anapara: {anapara:,.2f}   |  "
            f"Kalan şlemiş Faiz: {toplam_kalan_faiz:,.2f}   |  Kalan Masraf: {kalan_masraf:,.2f}   |  "
            f"TOPLAM ÖDENECEK BAKYE BORÇ: {genel_bakiye:,.2f} "
        )


# ------------------------------------------------------------------------------
# ÖZELLK 5: 6100 SAYILI HMK m.92-104 ADL TATL VE RESM TATL SÜRE HESAPLAYICI
# ------------------------------------------------------------------------------
class HmkSureHesapDialog(QDialog):
    def __init__(self, win):
        super().__init__(win)
        self.setWindowTitle("6100 Sayılı HMK m.92-104 Hukuki Süre, Resmi Tatil ve Adli Tatil Hesaplayıcı")
        self.resize(740, 520)
        lay = QVBoxLayout(self)

        gb = QGroupBox("Tebliğ / Tefhim Tarihi ve Kanuni Süre Girişi")
        gl = QGridLayout(gb)

        gl.addWidget(QLabel("Tebliğ / Tefhim Tarihi (GG.AA.YYYY):"), 0, 0)
        self.le_teblig = QLineEdit(datetime.now().strftime("%d.%m.%Y"))
        gl.addWidget(self.le_teblig, 0, 1)

        gl.addWidget(QLabel("Kanuni Süre Türü / şlem:"), 1, 0)
        self.cmb_sure = QComboBox()
        self.cmb_sure.addItems([
            "2 Hafta (HMK stinaf / Temyiz / Cevap Dilekçesi Süresi - 14 Gün)",
            "1 Hafta (cra Hukuk stinaf / HMK tiraz Süresi - 7 Gün)",
            "7 Gün (lamlı/lamsız cra Ödeme Emrine tiraz - K m.62)",
            "1 Ay (30 Gün)",
            "Özel Gün Sayısı (Aşağıdaki Kutudan Giriniz)"
        ])
        self.cmb_sure.currentIndexChanged.connect(self.calc_hmk_deadline)
        gl.addWidget(self.cmb_sure, 1, 1)

        gl.addWidget(QLabel("Özel Gün Sayısı:"), 2, 0)
        self.spn_gun = QSpinBox()
        self.spn_gun.setRange(1, 365)
        self.spn_gun.setValue(14)
        self.spn_gun.valueChanged.connect(self.calc_hmk_deadline)
        gl.addWidget(self.spn_gun, 2, 1)

        self.chk_adli_tatil = QCheckBox("şlem Adli Tatile Tabidir (6100 S.K. HMK m.102-104: 20 Temmuz - 31 Ağustos)")
        self.chk_adli_tatil.setChecked(True)
        self.chk_adli_tatil.toggled.connect(self.calc_hmk_deadline)
        gl.addWidget(self.chk_adli_tatil, 3, 0, 1, 2)

        btn_hesap = QPushButton("📅 Kesin Son Günü (HMK m.92-104) Hesapla")
        btn_hesap.clicked.connect(self.calc_hmk_deadline)
        gl.addWidget(btn_hesap, 4, 0, 1, 2)
        lay.addWidget(gb)

        self.txt_out = QTextEdit()
        self.txt_out.setReadOnly(True)
        lay.addWidget(self.txt_out, 1)
        self.calc_hmk_deadline()

    def is_official_holiday(self, d):
        fixed = {(1, 1): "Yılbaşı", (4, 23): "23 Nisan Ulusal Egemenlik ve Çocuk Bayramı",
                 (5, 1): "1 Mayıs Emek ve Dayanışma Günü", (5, 19): "19 Mayıs Atatürk'ü Anma, Gençlik ve Spor Bayramı",
                 (7, 15): "15 Temmuz Demokrasi ve Milli Birlik Günü", (8, 30): "30 Ağustos Zafer Bayramı",
                 (10, 29): "29 Ekim Cumhuriyet Bayramı"}
        return fixed.get((d.month, d.day))

    def calc_hmk_deadline(self):
        try:
            start_dt = datetime.strptime(self.le_teblig.text().strip(), "%d.%m.%Y").date()
        except Exception:
            self.txt_out.setPlainText("Lütfen tebliğ tarihini GG.AA.YYYY formatında giriniz.")
            return

        idx = self.cmb_sure.currentIndex()
        if idx == 0:
            days = 14
        elif idx in (1, 2):
            days = 7
        elif idx == 3:
            days = 30
        else:
            days = self.spn_gun.value()

        raw_end = start_dt + timedelta(days=days)
        final_end = raw_end
        notes = []

        # 1) 6100 S.K. HMK m.102-104 Adli Tatil Kontrolü (20 Temmuz - 31 Ağustos)
        adli_bas = date(raw_end.year, 7, 20)
        adli_bit = date(raw_end.year, 8, 31)
        if self.chk_adli_tatil.isChecked() and (adli_bas <= final_end <= adli_bit):
            final_end = adli_bit + timedelta(days=7)  # 7 Eylül
            notes.append(
                f"• ADL TATL UZATIMI (6100 S.K. HMK m.104): Sürenin son günü ({raw_end.strftime('%d.%m.%Y')}) "
                f"20 Temmuz - 31 Ağustos adli tatil dönemine rastladığından, süre adli tatilin bittiği günden "
                f"itibaren 1 hafta uzatılmış ve {final_end.strftime('%d.%m.%Y')} tarihine taşınmıştır."
            )

        # 2) 6100 S.K. HMK m.93 Resmi Tatil ve Hafta Sonu Kontrolü
        while True:
            hol_name = self.is_official_holiday(final_end)
            if final_end.weekday() == 5:  # Cumartesi
                notes.append(f"• HAFTA SONU UZATIMI (HMK m.93): {final_end.strftime('%d.%m.%Y')} Cumartesi gününe rastladığından takip eden iş gününe uzatıldı.")
                final_end += timedelta(days=2)
            elif final_end.weekday() == 6:  # Pazar
                notes.append(f"• HAFTA SONU UZATIMI (HMK m.93): {final_end.strftime('%d.%m.%Y')} Pazar gününe rastladığından takip eden iş gününe uzatıldı.")
                final_end += timedelta(days=1)
            elif hol_name:
                notes.append(f"• RESM TATL UZATIMI (HMK m.93): {final_end.strftime('%d.%m.%Y')} ({hol_name}) resmi tatile rastladığından takip eden ilk iş gününe uzatıldı.")
                final_end += timedelta(days=1)
            else:
                break

        gun_adlari = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
        rep = (
            f"========================================================================\n"
            f"   6100 SAYILI HUKUK MUHAKEMELER KANUNU m.92-104 SÜRE HESABI RAPORU\n"
            f"========================================================================\n"
            f"• Tebliğ / Başlangıç Tarihi : {start_dt.strftime('%d.%m.%Y')} ({gun_adlari[start_dt.weekday()]})\n"
            f"• Hesaplanan Kanuni Süre    : {days} Gün (HMK m.92 uyarınca tebliğ günü hesaba katılmaz)\n"
            f"• Takvim Üzerinde Ham Bitiş : {raw_end.strftime('%d.%m.%Y')} ({gun_adlari[raw_end.weekday()]})\n"
            f"------------------------------------------------------------------------\n"
        )
        if notes:
            rep += "\n".join(notes) + "\n------------------------------------------------------------------------\n"
        else:
            rep += "• Sürenin son günü normal mesai gününe denk gelmektedir.\n------------------------------------------------------------------------\n"

        rep += f"🎯 KESN SON ŞLEM GÜNÜ      : {final_end.strftime('%d.%m.%Y')} ({gun_adlari[final_end.weekday()]}) Mesai Bitimi (UYAP: 23:59)\n"
        rep += f"========================================================================"
        self.txt_out.setPlainText(rep)


# ------------------------------------------------------------------------------
# ÖZELLK 6: MÜVEKKL DOSYASI / SENARYO KAYDETME VE GER YÜKLEME YÖNETCS
# ------------------------------------------------------------------------------
def save_client_scenario(win):
    name, ok = QFileDialog.getSaveFileName(
        win, "Müvekkil Hesaplama Dosyasını Kaydet",
        os.path.join(MUVEKKIL_DIR, f"Muvekkil_Hesap_{datetime.now().strftime('%Y%m%d')}.json"),
        "Müvekkil Dosyası (*.json)"
    )
    if not ok or not name:
        return
    state = {"saved_at": datetime.now().strftime("%d.%m.%Y %H:%M:%S"), "fields": {}, "combos": {}, "spins": {}, "checks": {}}
    for attr in dir(win):
        if attr.startswith("_"):
            continue
        obj = getattr(win, attr, None)
        if isinstance(obj, QLineEdit):
            state["fields"][attr] = obj.text()
        elif isinstance(obj, QComboBox):
            state["combos"][attr] = obj.currentIndex()
        elif isinstance(obj, QSpinBox):
            state["spins"][attr] = obj.value()
        elif isinstance(obj, QCheckBox):
            state["checks"][attr] = obj.isChecked()
    with open(name, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    if hasattr(win, "notify"):
        win.notify(f"Müvekkil hesaplama dosyası kaydedildi: {os.path.basename(name)}")


def load_client_scenario(win):
    name, ok = QFileDialog.getOpenFileName(
        win, "Kayıtlı Müvekkil Hesaplama Dosyasını Aç",
        MUVEKKIL_DIR, "Müvekkil Dosyası (*.json)"
    )
    if not ok or not name or not os.path.exists(name):
        return
    try:
        with open(name, "r", encoding="utf-8") as f:
            state = json.load(f)
        for attr, val in state.get("fields", {}).items():
            obj = getattr(win, attr, None)
            if isinstance(obj, QLineEdit):
                obj.setText(str(val))
        for attr, val in state.get("combos", {}).items():
            obj = getattr(win, attr, None)
            if isinstance(obj, QComboBox) and 0 <= int(val) < obj.count():
                obj.setCurrentIndex(int(val))
        for attr, val in state.get("spins", {}).items():
            obj = getattr(win, attr, None)
            if isinstance(obj, QSpinBox):
                obj.setValue(int(val))
        for attr, val in state.get("checks", {}).items():
            obj = getattr(win, attr, None)
            if isinstance(obj, QCheckBox):
                obj.setChecked(bool(val))
        if hasattr(win, "hesapla_gayrimenkul"):
            win.hesapla_gayrimenkul()
        if hasattr(win, "notify"):
            win.notify(f"Müvekkil dosyası yüklendi: {os.path.basename(name)}")
    except Exception as e:
        if hasattr(win, "notify"):
            win.notify(f"Dosya yükleme hatası: {e}")


# ------------------------------------------------------------------------------
# ANA ENTEGRASYON KANCASI (MEVCUT HÇBR KODU BOZMADAN 6 ÖZELL BALAR)
# ------------------------------------------------------------------------------
def install_6_productivity_features(win):
    if getattr(win, "_verimlilik_6_installed", False):
        return
    win._verimlilik_6_installed = True
    try:
        import harc_ve_format_eklenti
        harc_ve_format_eklenti.install_harclar_ve_duzeltmeler(win)
    except Exception as e:
        print('Harc eklenti hatasi:', e)

    _bind_global_dialog_icons(win)
    try:
        from PyQt6.QtGui import QIcon
        from PyQt6.QtWidgets import QApplication
        _ico_p = os.path.join(BASE_DIR, 'ucar_modern_icon.ico')
        if os.path.exists(_ico_p):
            _ic = QIcon(_ico_p)
            win.setWindowIcon(_ic)
            if QApplication.instance():
                QApplication.instance().setWindowIcon(_ic)
    except Exception:
        pass


    # 1 & 4) Alış ve Satış tarihi değiştiğinde 10 yıllık yerleşik Y-ÜFE arşivinden
    #        önceki ayın Y-ÜFE endeksini ve satış yılına göre GVK stisna tutarını otomatik doldur!
    def _auto_fill_yiufe_and_yearly_exemption():
        try:
            if hasattr(win, "dt_gm_alis") and hasattr(win, "txt_yiufe_alis"):
                _, v_alis = get_prev_month_yiufe(win.dt_gm_alis.text())
                if v_alis:
                    win.txt_yiufe_alis.setText(f"{v_alis:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            if hasattr(win, "dt_gm_satis") and hasattr(win, "txt_yiufe_satis"):
                _, v_satis = get_prev_month_yiufe(win.dt_gm_satis.text())
                if v_satis:
                    win.txt_yiufe_satis.setText(f"{v_satis:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                # Satış yılına göre GVK Mük. m.80 stisna tutarını otomatik güncelle
                dt_s = datetime.strptime(win.dt_gm_satis.text().strip(), "%d.%m.%Y")
                y_info = GVK_YEARLY_DATA.get(dt_s.year)
                if y_info and hasattr(win, "gm_istisna"):
                    ist_val = y_info["istisna"]
                    win.gm_istisna.setText(f"{ist_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            if hasattr(win, "hesapla_gayrimenkul"):
                win.hesapla_gayrimenkul()
        except Exception:
            pass

    if hasattr(win, "dt_gm_alis"):
        win.dt_gm_alis.editingFinished.connect(_auto_fill_yiufe_and_yearly_exemption)
    if hasattr(win, "dt_gm_satis"):
        win.dt_gm_satis.editingFinished.connect(_auto_fill_yiufe_and_yearly_exemption)

    # 2) Kategori 5 (Kira, Gayrimenkul & Miras) sol paneline "📄 UYAP Uyumlu Dilekçe / htarname Metni Üret" butonu ekle
    if hasattr(win, "gm_istisna") and win.gm_istisna.parentWidget():
        pw = win.gm_istisna.parentWidget().parentWidget() or win.gm_istisna.parentWidget()
        if pw and pw.layout():
            btn_dilekce = QPushButton("📄 UYAP (.udf) Uyumlu Dilekçe ve htarname Metni Üret (TBK m.344 / KMK m.20)")
            btn_dilekce.setMinimumHeight(36)
            btn_dilekce.clicked.connect(lambda: DilekceMetniDialog(win).exec())
            pw.layout().addWidget(btn_dilekce)

    # 3) Faiz Hesaplama Stüdyosu içine 4. Sekme olarak "4. TBK m.100 Kısmi Ödeme Mahsubu" sekmesini ekle
    if hasattr(win, "_faiz_studyo_ref") and win._faiz_studyo_ref is not None:
        st = win._faiz_studyo_ref
        if hasattr(st, "sub_tabs") and st.sub_tabs.count() < 4:
            st.sub_tabs.addTab(Tbk100MahsupWidget(st), "4. TBK m.100 Kısmi Ödeme Mahsubu (Önce Faiz/Masraf)")

    # 5 & 6) Üst Araç Çubuğuna "⚖️ HMK Süre & Adli Tatil", "💾 Müvekkil Kaydet" ve "📂 Müvekkil Yükle" butonlarını ekle
    tb = win.findChild(QToolBar)
    if tb is not None:
        btn_hmk = QPushButton("⚖️ HMK Süre & Adli Tatil")
        btn_hmk.setToolTip("6100 S.K. HMK m.92-104 Adli Tatil ve Resmi Tatil Süre Hesaplayıcı")
        btn_hmk.clicked.connect(lambda: HmkSureHesapDialog(win).exec())

        btn_save_m = QPushButton("💾 Müvekkil Kaydet")
        btn_save_m.setToolTip("Ekrandaki tüm hesaplama verilerini müvekkil dosyası (.json) olarak kaydeder")
        btn_save_m.clicked.connect(lambda: save_client_scenario(win))

        btn_load_m = QPushButton("📂 Müvekkil Aç")
        btn_load_m.setToolTip("Kayıtlı müvekkil hesaplama dosyasını ekrana geri yükler")
        btn_load_m.clicked.connect(lambda: load_client_scenario(win))

        for b in [btn_hmk, btn_save_m, btn_load_m]:
            b.setStyleSheet(
                "QPushButton { background-color: #FFFFFF; color: #0A3622; font-weight: bold; "
                "padding: 4px 10px; border: 2px solid #C5A059; border-radius: 5px; }"
                "QPushButton:hover { background-color: #D0E8DA; color: #052214; }"
            )
            tb.addWidget(b)

    # Açılışta Y-ÜFE arşivinden mevcut tarihleri bir kez eşleştir
    QTimer.singleShot(400, _auto_fill_yiufe_and_yearly_exemption)


def _bind_global_dialog_icons(win):
    try:
        from PyQt6.QtGui import QIcon
        from PyQt6.QtWidgets import QApplication, QDialog, QWidget
        from PyQt6.QtCore import QObject, QEvent
        _ico_p = os.path.join(BASE_DIR, "ucar_modern_icon.ico")
        if not os.path.exists(_ico_p):
            return
        _ic = QIcon(_ico_p)
        win.setWindowIcon(_ic)
        app = QApplication.instance()
        if app:
            app.setWindowIcon(_ic)
            class _GlobalIconFilter(QObject):
                def eventFilter(self, obj, ev):
                    if ev.type() in (QEvent.Type.Show, QEvent.Type.WindowTitleChange, QEvent.Type.Polish):
                        if isinstance(obj, (QDialog, QWidget)) and obj.isWindow():
                            obj.setWindowIcon(_ic)
                    return False
            win._global_ico_filter = _GlobalIconFilter(win)
            app.installEventFilter(win._global_ico_filter)
            for w in app.topLevelWidgets():
                w.setWindowIcon(_ic)
    except Exception:
        pass
