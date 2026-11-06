# -*- coding: utf-8 -*-
import os, sys, re, json

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QSpinBox,
    QTabWidget, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QTextEdit, QFileDialog, QMessageBox, QFrame
)
from PyQt6.QtCore import Qt, QObject, QEvent, QTimer, QDate
from PyQt6.QtGui import QFont, QIcon, QColor

# === 2026 GUNCEL MEVZUAT KATSAYILari VE TARIFE TABLOSU ===
MEVZUAT_2026 = {
    "icra": {
        "basvuru_harci": 690.20,
        "vekalet_suret_harci": 98.40,
        "pesin_harc_orani": 0.005,          # Binde 5 (%0,5)
        "tahsil_haciz_oncesi": 0.0455,      # %4,55
        "tahsil_haciz_sonrasi": 0.0910,     # %9,10
        "tahsil_satis_sonrasi": 0.1138,     # %11,38
        "cezaevi_harci_orani": 0.02,        # %2 (Alacaklidan)
        "tahliye_teslim_orani": 0.0227,     # %2,27
        "icra_inkar_asgari": 0.20           # %20 (IIK m.67)
    },
    "mahkeme": {
        "sulh_basvuru": 318.30,
        "asliye_basvuru": 690.20,
        "idare_basvuru": 690.20,
        "istinaf_basvuru": 1915.40,
        "temyiz_basvuru": 3296.80,
        "nispi_karar_ilam_orani": 0.06831,  # Binde 68,31 (%6,831)
        "pesin_harc_katsayisi": 0.25,       # 1/4 (%1,70775)
        "Discovery_harci": 4882.70,         # Kesif Harci
        "tebligat_birim": 210.00,           # HMK Yonetmeligi 1 Adet Tebligat
        "gider_avansi_sabit": 940.00,       # Yonetmelik diger is ve islemler
        "tanik_birim": 250.00,
        "bilirkisi_orta": 5500.00
    }
}

def format_tr_number(val, decimals=2):
    try:
        s = f"{float(val):,.{decimals}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"

def parse_tr_number(text):
    if not text:
        return 0.0
    s = str(text).strip().replace("TL", "").replace("", "").replace("%", "").strip()
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
    try:
        return float(s)
    except Exception:
        return 0.0

def live_thousand_separator(line_edit: QLineEdit):
    if getattr(line_edit, "_thousand_bound", False):
        return
    line_edit._thousand_bound = True

    def _on_text_edited(text):
        if getattr(line_edit, "_formatting_lock", False):
            return
        # Sadece rakam, nokta, virgul ve eksi iceriyorsa formatla (tarih veya metin kutularina dokunma)
        clean_check = text.strip()
        if not clean_check:
            return
        if re.search(r"[a-zA-ZğüşıöçÜŞÖÇ/\\:]", clean_check):
            return
        # Eger tarih formatiysa (GG.AA.YYYY gibi 2 nokta ve sonu 4 hane) dokunma
        if re.match(r"^\d{1,2}\.\d{1,2}\.\d{4}$", clean_check):
            return

        cursor_pos = line_edit.cursorPosition()
        digits_before = len(re.sub(r"[^\d,]", "", clean_check[:cursor_pos]))

        raw = re.sub(r"[^\d,]", "", clean_check)
        if not raw:
            return

        if "," in raw:
            int_part, dec_part = raw.split(",", 1)
            dec_part = re.sub(r"[^\d]", "", dec_part)[:2]
            has_comma = True
        else:
            int_part, dec_part = raw, ""
            has_comma = False

        int_part = int_part.lstrip("0") or "0"
        # Binlik noktalama
        rev = int_part[::-1]
        grouped = ".".join(rev[i:i+3] for i in range(0, len(rev), 3))[::-1]
        new_text = f"{grouped},{dec_part}" if has_comma else grouped

        if new_text != text:
            line_edit._formatting_lock = True
            line_edit.setText(new_text)
            # Imlec konumunu koru
            new_pos = 0
            seen = 0
            for ch in new_text:
                if seen >= digits_before:
                    break
                new_pos += 1
                if ch.isdigit() or ch == ",":
                    seen += 1
            line_edit.setCursorPosition(new_pos)
            line_edit._formatting_lock = False

    line_edit.textEdited.connect(_on_text_edited)


class _GlobalInputAndFontFilter(QObject):
    def eventFilter(self, obj, ev):
        if ev.type() in (QEvent.Type.Show, QEvent.Type.ChildAdded, QEvent.Type.Polish):
            if isinstance(obj, QLineEdit):
                live_thousand_separator(obj)
        return False


class IcraHarclariWidget(QWidget):
    """492 Sayili Harclar Kanunu ve IIK Kapsaminda Icra Harclari ve Tahsilat Hesaplayici"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        lbl_head = QLabel("⚖️ CRA HARÇLARI, TAHSL HARCI VE FERAGAT MAHSUBU (492 S.K. & K)")
        lbl_head.setStyleSheet("font-size: 14pt; font-weight: bold; color: #0F3826; padding: 4px 0px;")
        lay.addWidget(lbl_head)

        top_box = QGroupBox("Takip ve Tahsilat Parametreleri")
        grid = QGridLayout(top_box)
        grid.setVerticalSpacing(10)
        grid.setHorizontalSpacing(14)

        grid.addWidget(QLabel("Takip Türü:"), 0, 0)
        self.cmb_takip_turu = QComboBox()
        self.cmb_takip_turu.addItems([
            "lamsız cra Takibi (Genel Haciz / Kambiyo)",
            "lamlı cra Takibi (Mahkeme Kararına Dayalı)",
            "Kira Alacağı ve Tahliye Talepli Takip",
            "Rehnin / poteğin Paraya Çevrilmesi"
        ])
        grid.addWidget(self.cmb_takip_turu, 0, 1)

        grid.addWidget(QLabel("Takip Çıkış Tutarı (TL):"), 0, 2)
        self.inp_takip_tutari = QLineEdit("250.000")
        live_thousand_separator(self.inp_takip_tutari)
        grid.addWidget(self.inp_takip_tutari, 0, 3)

        grid.addWidget(QLabel("Tahsilat Aşaması (Yönetmelik/Tarife):"), 1, 0)
        self.cmb_asama = QComboBox()
        self.cmb_asama.addItems([
            "1. Ödeme/cra Emri Tebliği Üzerine - Hacizden Önce (%4,55)",
            "2. Hacizden Sonra ve Satıştan Önce Ödeme (%9,10)",
            "3. Haczedilen Malın Satılıp Paraya Çevrilmesiyle (%11,38)"
        ])
        grid.addWidget(self.cmb_asama, 1, 1)

        grid.addWidget(QLabel("şlemiş Faiz ve Masraf (TL):"), 1, 2)
        self.inp_ek_faiz_masraf = QLineEdit("45.000")
        live_thousand_separator(self.inp_ek_faiz_masraf)
        grid.addWidget(self.inp_ek_faiz_masraf, 1, 3)

        self.chk_feragat = QCheckBox("Takipten Feragat / Haricen Tahsil (Harçlar K. m.23 uyarınca 1/3 Harç)")
        self.chk_cezaevi = QCheckBox("Cezaevi Yapı Harcı (%2) Hesapla (Kamuya ait muafiyet yoksa)")
        self.chk_cezaevi.setChecked(True)
        self.chk_vekalet = QCheckBox("AAÜT cra Vekalet Ücreti Hesapla")
        self.chk_vekalet.setChecked(True)

        grid.addWidget(self.chk_feragat, 2, 0, 1, 2)
        grid.addWidget(self.chk_cezaevi, 2, 2, 1, 1)
        grid.addWidget(self.chk_vekalet, 2, 3, 1, 1)

        btn_calc = QPushButton("🔍 CRA HARÇLARINI VE KAPANIŞ BAKYESN HESAPLA")
        btn_calc.setMinimumHeight(38)
        btn_calc.clicked.connect(self.hesapla)
        grid.addWidget(btn_calc, 3, 0, 1, 4)

        lay.addWidget(top_box)

        self.tbl = QTableWidget(0, 4)
        self.tbl.setHorizontalHeaderLabels(["Harç / Kalem Adı", "Yasal Dayanak & Oran", "Tutar (TL)", "Yükümlü / Mahsup Notu"])
        self.tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.tbl.setAlternatingRowColors(True)
        lay.addWidget(self.tbl)

        self.lbl_ozet = QLabel()
        self.lbl_ozet.setStyleSheet("background-color: #FFFFFF; border: 2px solid #0F3826; border-radius: 6px; padding: 10px; font-size: 11pt; font-weight: bold; color: #2C1A12;")
        self.lbl_ozet.setWordWrap(True)
        lay.addWidget(self.lbl_ozet)

        self.hesapla()

    def _calculate_icra_aaut(self, matrah):
        # 2026 AAUT Icra Takipleri Kademeli Hesap (Ilk 400.000 TL %16 vb.)
        dilimler = [
            (400000, 0.16),
            (400000, 0.15),
            (800000, 0.14),
            (1200000, 0.11),
            (1600000, 0.08),
            (2000000, 0.05),
            (2400000, 0.03),
            (2800000, 0.02),
        ]
        kalan = matrah
        ucret = 0.0
        for limit, oran in dilimler:
            if kalan <= 0:
                break
            taban = min(kalan, limit)
            ucret += taban * oran
            kalan -= taban
        if kalan > 0:
            ucret += kalan * 0.01
        maktu = 6000.0
        if matrah <= maktu:
            return matrah
        return max(ucret, maktu)

    def hesapla(self):
        cfg = MEVZUAT_2026["icra"]
        takip_turu = self.cmb_takip_turu.currentIndex()
        asil_alacak = parse_tr_number(self.inp_takip_tutari.text())
        ek_tutar = parse_tr_number(self.inp_ek_faiz_masraf.text())
        toplam_matrah = asil_alacak + ek_tutar

        ilamli_mi = (takip_turu == 1)
        basvuru = cfg["basvuru_harci"]
        vekalet_harc = cfg["vekalet_suret_harci"]
        pesin_harc = 0.0 if ilamli_mi else (asil_alacak * cfg["pesin_harc_orani"])

        asama_idx = self.cmb_asama.currentIndex()
        if asama_idx == 0:
            oran = cfg["tahsil_haciz_oncesi"]
            asama_ad = "Hacizden Önce (%4,55)"
        elif asama_idx == 1:
            oran = cfg["tahsil_haciz_sonrasi"]
            asama_ad = "Hacizden Sonra Satıştan Önce (%9,10)"
        else:
            oran = cfg["tahsil_satis_sonrasi"]
            asama_ad = "Satış Sonrası (%11,38)"

        if self.chk_feragat.isChecked():
            oran = oran / 3.0
            asama_ad += " [Feragat 1/3 Oranı]"

        brut_tahsil_harci = toplam_matrah * oran
        net_tahsil_harci = max(0.0, brut_tahsil_harci - pesin_harc)
        cezaevi_harci = (toplam_matrah * cfg["cezaevi_harci_orani"]) if self.chk_cezaevi.isChecked() else 0.0
        icra_vekalet = self._calculate_icra_aaut(asil_alacak) if self.chk_vekalet.isChecked() else 0.0

        rows = [
            ("cra Başvurma Harcı", "492 S.K. (1) Sayılı Tarife (Maktu)", format_tr_number(basvuru), "Takip açılışında alacaklı yatırır, borçluya yüklenir"),
            ("Peşin Harç", "lamlı Takipte Alınmaz" if ilamli_mi else "Harçlar K. m.29 (Binde 5 - %0,5)", format_tr_number(pesin_harc), "Tahsil harcından mahsup edilir"),
            ("Vekalet Suret Harcı", "492 S.K. Maktu Vekalet Harcı", format_tr_number(vekalet_harc), "Takip açılış masrafı"),
            ("Brüt Tahsil Harcı", f"492 S.K. - {asama_ad}", format_tr_number(brut_tahsil_harci), f"Matrah: {format_tr_number(toplam_matrah)} TL üzerinden"),
            ("Peşin Harç Mahsubu Sonrası Net Tahsil Harcı", "Brüt Tahsil Harcı (-) Peşin Harç", format_tr_number(net_tahsil_harci), "Borçluya ait (cra veznesine ödenecek net tahsil harcı)"),
            ("Cezaevi Yapı Harcı", "2548 S.K. m.1 (%2)", format_tr_number(cezaevi_harci), "Alacaklıya ödenen paradan kesilir (Borçluya yüklenmez)"),
            ("AAÜT cra Vekalet Ücreti", "2026 AAÜT cra Kademeli Tarife", format_tr_number(icra_vekalet), "Borçluya ait kapak hesabı kalemi")
        ]

        self.tbl.setRowCount(len(rows))
        for r, (k, d, t, n) in enumerate(rows):
            self.tbl.setItem(r, 0, QTableWidgetItem(k))
            self.tbl.setItem(r, 1, QTableWidgetItem(d))
            it_t = QTableWidgetItem(f"{t} TL")
            it_t.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it_t.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            self.tbl.setItem(r, 2, it_t)
            self.tbl.setItem(r, 3, QTableWidgetItem(n))

        acilis_toplam = basvuru + pesin_harc + vekalet_harc
        kapak_toplam = toplam_matrah + net_tahsil_harci + icra_vekalet + basvuru + vekalet_harc
        alacakli_net = toplam_matrah + icra_vekalet - cezaevi_harci

        self.lbl_ozet.setText(
            f"📌 TAKP AÇILIŞINDA YATIRILACAK TOPLAM HARÇ: {format_tr_number(acilis_toplam)} TL   |   "
            f"📌 DOSYA KAPANIŞ (BORÇLU TOPLAM YÜKÜ): {format_tr_number(kapak_toplam)} TL   |   "
            f"📌 ALACAKLIYA (VEKLNE) NET ÖDENECEK: {format_tr_number(alacakli_net)} TL"
        )


class MahkemeHarclariWidget(QWidget):
    """492 Sayili Harclar Kanunu ve HMK Gider/Delil Avansi Yonetmeligi Hesaplayici"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        lbl_head = QLabel("🏛️ MAHKEME HARÇLARI, HMK GDER AVANSI VE DELL AVANSI (MEVZUAT & YÖNETMELK UYUMLU)")
        lbl_head.setStyleSheet("font-size: 14pt; font-weight: bold; color: #0F3826; padding: 4px 0px;")
        lay.addWidget(lbl_head)

        box = QGroupBox("Mahkeme, Dava Değeri ve Yönetmelik Parametreleri")
        grid = QGridLayout(box)
        grid.setVerticalSpacing(10)
        grid.setHorizontalSpacing(14)

        grid.addWidget(QLabel("Mahkeme / Yargı Kolu:"), 0, 0)
        self.cmb_mahkeme = QComboBox()
        self.cmb_mahkeme.addItems([
            "Asliye Hukuk / Asliye Ticaret / ş / Aile Mahkemesi",
            "Sulh Hukuk Mahkemesi (Kira, Tahliye, Ortaklığın Giderilmesi)",
            "Tüketici Mahkemesi (Tüketici Harçtan Muaf - 6502 S.K.)",
            "dare ve Vergi Mahkemeleri (2577 S.K.)",
            "Bölge Adliye Mahkemesi (stinaf Başvurusu)",
            "Yargıtay / Danıştay (Temyiz Başvurusu)"
        ])
        grid.addWidget(self.cmb_mahkeme, 0, 1)

        grid.addWidget(QLabel("Dava / Harca Esas Değer (TL):"), 0, 2)
        self.inp_dava_degeri = QLineEdit("1.250.000")
        live_thousand_separator(self.inp_dava_degeri)
        grid.addWidget(self.inp_dava_degeri, 0, 3)

        grid.addWidget(QLabel("Dava Niteliği:"), 1, 0)
        self.cmb_nitelik = QComboBox()
        self.cmb_nitelik.addItems([
            "Nispi Harca Tabi Dava (Alacak, Tazminat, Tapu ptal, Kira Yıllık)",
            "Maktu Harca Tabi Dava (Tespit, Tahliye-Kirasız, Boşanma)",
            "cra Takibine tirazın ptali Davası (Peşin Harç Mahsuplu)"
        ])
        grid.addWidget(self.cmb_nitelik, 1, 1)

        grid.addWidget(QLabel("Taraf Sayısı (HMK Gider Avansı):"), 1, 2)
        self.spn_taraf = QSpinBox()
        self.spn_taraf.setRange(2, 50)
        self.spn_taraf.setValue(2)
        grid.addWidget(self.spn_taraf, 1, 3)

        self.chk_bilirkisi = QCheckBox("Bilirkişi Delil Avansı Ekle (Yönetmelik Tarifesi)")
        self.chk_bilirkisi.setChecked(True)
        self.chk_kesif = QCheckBox("Keşif Harcı ve Keşif Avansı Ekle")
        self.spn_tanik = QSpinBox()
        self.spn_tanik.setRange(0, 20)
        self.spn_tanik.setValue(2)
        self.spn_tanik.setPrefix("Tanık Sayısı: ")

        grid.addWidget(self.chk_bilirkisi, 2, 0, 1, 2)
        grid.addWidget(self.chk_kesif, 2, 2, 1, 1)
        grid.addWidget(self.spn_tanik, 2, 3, 1, 1)

        btn = QPushButton("⚖️ MAHKEME HARÇLARINI VE YÖNETMELK GDER AVANSINI HESAPLA")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self.hesapla)
        grid.addWidget(btn, 3, 0, 1, 4)

        lay.addWidget(box)

        self.tbl = QTableWidget(0, 4)
        self.tbl.setHorizontalHeaderLabels(["Harç / Avans Kalemi", "Mevzuat & Yönetmelik Dayanağı", "Tutar (TL)", "Açıklama"])
        self.tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.tbl.setAlternatingRowColors(True)
        lay.addWidget(self.tbl)

        self.lbl_ozet = QLabel()
        self.lbl_ozet.setStyleSheet("background-color: #FFFFFF; border: 2px solid #0F3826; border-radius: 6px; padding: 10px; font-size: 11pt; font-weight: bold; color: #2C1A12;")
        self.lbl_ozet.setWordWrap(True)
        lay.addWidget(self.lbl_ozet)

        self.hesapla()

    def hesapla(self):
        cfg = MEVZUAT_2026["mahkeme"]
        m_idx = self.cmb_mahkeme.currentIndex()
        n_idx = self.cmb_nitelik.currentIndex()
        deger = parse_tr_number(self.inp_dava_degeri.text())
        taraf_sayisi = self.spn_taraf.value()
        tanik_sayisi = self.spn_tanik.value()

        # 1. Basvurma Harci
        if m_idx == 1:
            basvuru = cfg["sulh_basvuru"]
        elif m_idx == 2:
            basvuru = 0.0
        elif m_idx == 4:
            basvuru = cfg["istinaf_basvuru"]
        elif m_idx == 5:
            basvuru = cfg["temyiz_basvuru"]
        else:
            basvuru = cfg["asliye_basvuru"]

        # 2. Karar ve Ilam Harci (Binde 68,31) ve 1/4 Pesin Harc
        if m_idx == 2:
            toplam_karar_harci = 0.0
            pesin_harc = 0.0
            bakiye_karar_harci = 0.0
            harc_aciklama = "6502 S.K. m.73/2 uyarınca tüketici davaları harçtan muaftır"
        elif n_idx == 1 or m_idx in (4, 5):
            # Maktu Karar ve Ilam Harci
            toplam_karar_harci = basvuru
            pesin_harc = basvuru
            bakiye_karar_harci = 0.0
            harc_aciklama = "Maktu Karar ve lam Harcı"
        else:
            toplam_karar_harci = max(basvuru, deger * cfg["nispi_karar_ilam_orani"])
            pesin_harc = toplam_karar_harci * cfg["pesin_harc_katsayisi"]
            if n_idx == 2:
                # Itirazin Iptali: Icra takibindeki binde 5 pesin harc mahsup edilir
                icra_pesin = deger * 0.005
                pesin_harc = max(0.0, pesin_harc - icra_pesin)
                harc_aciklama = f"tirazın ptali: crada yatırılan {format_tr_number(icra_pesin)} TL peşin harç mahsup edildi"
            else:
                harc_aciklama = "Nispi Karar ve lam Harcının (Binde 68,31) 1/4'ü dava açılışında peşin alınır"
            bakiye_karar_harci = max(0.0, toplam_karar_harci - (toplam_karar_harci * 0.25))

        vekalet_suret = 98.40 if m_idx != 2 else 0.0

        # 3. HMK Gider Avansi Tarifesi (Yonetmelik: Taraf Sayisi x 5 kat tebligat + diger is ve islemler)
        zorunlu_gider_avansi = (taraf_sayisi * 5 * cfg["tebligat_birim"]) + cfg["gider_avansi_sabit"]

        # 4. HMK m.324 Delil Avansi (Bilirkisi + Kesif + Tanik)
        bilirkisi_ucreti = cfg["bilirkisi_orta"] if self.chk_bilirkisi.isChecked() else 0.0
        kesif_harci = cfg["Discovery_harci"] if self.chk_kesif.isChecked() else 0.0
        tanik_avansi = tanik_sayisi * (cfg["tanik_birim"] + cfg["tebligat_birim"])
        delil_avansi_toplam = bilirkisi_ucreti + kesif_harci + tanik_avansi

        rows = [
            ("Başvurma Harcı", "492 S.K. (1) Sayılı Tarife A/I", format_tr_number(basvuru), "Dava açılışında maktu alınır"),
            ("Peşin Karar ve lam Harcı (1/4)", "492 S.K. m.28 (Binde 68,31'in 1/4'ü = %1,70775)", format_tr_number(pesin_harc), harc_aciklama),
            ("Vekaletname Suret Harcı", "492 S.K. Maktu", format_tr_number(vekalet_suret), "Avukatla takip edilen davalarda"),
            ("HMK Zorunlu Gider Avansı (Yönetmelik)", f"HMK Gider Avansı Tar. ({taraf_sayisi} Taraf x 5 Tebligat + Sabit)", format_tr_number(zorunlu_gider_avansi), "HMK m.114/1-g dava şartı (2 hafta kesin süre)"),
            ("HMK m.324 Delil Avansı (Bilirkişi/Keşif/Tanık)", "Bilirkişi + Keşif Harcı + Tanık Ücret/Tebligat", format_tr_number(delil_avansi_toplam), "Delil ikamesi için mahkeme veznesine yatırılır"),
            ("Toplam Nispi Karar ve lam Harcı (1/1)", "492 S.K. (Binde 68,31 - %6,831)", format_tr_number(toplam_karar_harci), "Hüküm kurulduğunda hesaplanan toplam karar harcı"),
            ("Karar Sonrası Bakiye Karar Harcı (3/4)", "492 S.K. m.28/1-a (%5,12325)", format_tr_number(bakiye_karar_harci), "Kararın kesinleşmesi/tebliği aşamasında haksız çıkan taraftan tahsil edilir")
        ]

        self.tbl.setRowCount(len(rows))
        for r, (k, d, t, n) in enumerate(rows):
            self.tbl.setItem(r, 0, QTableWidgetItem(k))
            self.tbl.setItem(r, 1, QTableWidgetItem(d))
            it_t = QTableWidgetItem(f"{t} TL")
            it_t.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it_t.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            self.tbl.setItem(r, 2, it_t)
            self.tbl.setItem(r, 3, QTableWidgetItem(n))

        dava_acilis_vezne = basvuru + pesin_harc + vekalet_suret + zorunlu_gider_avansi
        delil_dahil_toplam = dava_acilis_vezne + delil_avansi_toplam

        self.lbl_ozet.setText(
            f"🏛️ DAVA AÇILIŞINDA VEZNEYE YATIRILACAK ZORUNLU TOPLAM (HARÇ + GDER AVANSI): {format_tr_number(dava_acilis_vezne)} TL\n"
            f"📌 DELL AVANSI DAHL TOPLAM BAŞLANGIÇ MALYET: {format_tr_number(delil_dahil_toplam)} TL   |   "
            f"📌 KARAR SONRASI 3/4 BAKYE HARÇ: {format_tr_number(bakiye_karar_harci)} TL"
        )


def install_harclar_ve_duzeltmeler(win):
    if getattr(win, "_harclar_2026_installed", False):
        return
    win._harclar_2026_installed = True

    app = QApplication.instance()
    if app:
        # 1. Turkce I/i karakterlerinin tam ve kesintisiz okunmasi icin OpenType font ve yuksek paddingli stil
        app.setFont(QFont("Segoe UI", 10))
        qss_fix = """
        QWidget {
            font-family: "Segoe UI", "Calibri", "Arial", sans-serif;
        }
        QLabel, QCheckBox, QRadioButton {
            padding-top: 2px;
            padding-bottom: 2px;
        }
        QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
            min-height: 26px;
            padding: 3px 6px;
            font-size: 10.5pt;
        }
        QPushButton {
            min-height: 28px;
            padding: 5px 12px;
        }
        QHeaderView::section {
            padding: 6px;
            min-height: 24px;
        }
        """
        app.setStyleSheet(app.styleSheet() + "\n" + qss_fix)

        # 2. Tum QLineEdit kutularina canli binlik nokta (1.250.000) filtresi bagla
        win._global_input_filter = _GlobalInputAndFontFilter(win)
        app.installEventFilter(win._global_input_filter)
        for le in win.findChildren(QLineEdit):
            live_thousand_separator(le)

    # 3. Icra Harclari ve Mahkeme Harclari sekmelerini ana TabWidget icine yerlestir
    tabs = win.findChildren(QTabWidget)
    if tabs:
        main_tab = tabs[0]
        # Eger eski bir Mahkeme Harci veya Kira sekmesi varsa guncelle / temizle
        for i in range(main_tab.count() - 1, -1, -1):
            t_title = main_tab.tabText(i)
            if "cra Harç" in t_title or "Mahkeme Harç" in t_title:
                main_tab.removeTab(i)

        icra_w = IcraHarclariWidget(win)
        mahkeme_w = MahkemeHarclariWidget(win)
        main_tab.addTab(icra_w, "⚖️ cra Harçları Hesapla")
        main_tab.addTab(mahkeme_w, "🏛️ Mahkeme Harçları & Gider Avansı")