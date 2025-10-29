# Modern Stock Management System
import sys
import sqlite3
import hashlib
import csv
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QMessageBox,
                             QTableWidget, QTableWidgetItem, QMenu, QAction,
                             QInputDialog, QHeaderView, QFrame, QMainWindow,
                             QDialog, QFormLayout, QDialogButtonBox,
                             QAbstractItemView, QFileDialog, QStyle,
                             QMenuBar, QCheckBox)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QFont, QColor


# Sayısal Sıralama için Yardımcı Sınıf
class NumericTableWidgetItem(QTableWidgetItem):
    """Sadece sayısal olarak sıralama yapabilen özel bir tablo hücresi."""

    def __init__(self, display_text, sort_key):
        super().__init__(display_text)
        self.sort_key = sort_key  # Sıralama için kullanılacak ham sayısal değer

    def __lt__(self, other):
        # Diğer hücre de bu sınıftan ise, sayısal değerleri karşılaştır
        if isinstance(other, NumericTableWidgetItem):
            return self.sort_key < other.sort_key

        # Değilse, varsayılan metin karşılaştırmasını yap
        return super().__lt__(other)


# Database Manager
class VeritabaniYoneticisi:
    def __init__(self, db_adi="stok_veritabani.db"):
        self.baglanti = sqlite3.connect(db_adi)
        self.cursor = self.baglanti.cursor()
        self.veritabani_migrasyonu_kontrol_et()
        self.tablolari_olustur()

    def veritabani_migrasyonu_kontrol_et(self):
        """Check and update database schema"""
        try:
            self.cursor.execute("PRAGMA table_info(urunler)")
            mevcut_sutunlar = [row[1] for row in self.cursor.fetchall()]

            if 'min_stok' not in mevcut_sutunlar:
                print("VERİTABANI MİGRASYONU: 'min_stok' sütunu 'urunler' tablosuna ekleniyor...")
                self.cursor.execute("ALTER TABLE urunler ADD COLUMN min_stok INTEGER NOT NULL DEFAULT 10")
                self.baglanti.commit()
                print("Migrasyon tamamlandı.")

        except sqlite3.OperationalError as e:
            if "no such table: urunler" in str(e):
                pass
            else:
                raise e

    def tablolari_olustur(self):
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS urunler (
                id INTEGER PRIMARY KEY, 
                ad TEXT NOT NULL UNIQUE, 
                kategori TEXT, 
                fiyat REAL NOT NULL DEFAULT 0.0, 
                miktar INTEGER NOT NULL,
                min_stok INTEGER NOT NULL DEFAULT 10
            )""")
        self.cursor.execute(
            "CREATE TABLE IF NOT EXISTS kullanicilar (id INTEGER PRIMARY KEY, kullanici_adi TEXT NOT NULL UNIQUE, sifre_hash TEXT NOT NULL)")
        self.baglanti.commit()

    def urunleri_getir(self):
        self.cursor.execute("SELECT id, ad, kategori, fiyat, miktar, min_stok FROM urunler ORDER BY ad ASC")
        return self.cursor.fetchall()

    def urun_detay_getir(self, urun_id):
        self.cursor.execute("SELECT id, ad, kategori, fiyat, miktar, min_stok FROM urunler WHERE id = ?", (urun_id,))
        return self.cursor.fetchone()

    def urun_ekle(self, ad, kategori, fiyat, miktar, min_stok):
        try:
            self.cursor.execute("INSERT INTO urunler (ad, kategori, fiyat, miktar, min_stok) VALUES (?, ?, ?, ?, ?)",
                                (ad, kategori, fiyat, miktar, min_stok))
        except sqlite3.IntegrityError:
            self.cursor.execute("UPDATE urunler SET miktar = miktar + ? WHERE ad = ?", (miktar, ad))
        self.baglanti.commit()

    def urun_detay_guncelle(self, urun_id, yeni_ad, kategori, fiyat, min_stok):
        self.cursor.execute("UPDATE urunler SET ad = ?, kategori = ?, fiyat = ?, min_stok = ? WHERE id = ?",
                            (yeni_ad, kategori, fiyat, min_stok, urun_id))
        self.baglanti.commit()

    def urun_miktar_guncelle(self, urun_id, miktar_farki):
        self.cursor.execute("UPDATE urunler SET miktar = miktar + ? WHERE id = ?",
                            (miktar_farki, urun_id))
        self.baglanti.commit()

    def urun_sil(self, urun_id):
        self.cursor.execute("DELETE FROM urunler WHERE id = ?", (urun_id,))
        self.baglanti.commit()

    def mevcut_miktar_getir(self, urun_id):
        self.cursor.execute("SELECT miktar FROM urunler WHERE id = ?", (urun_id,))
        sonuc = self.cursor.fetchone()
        return sonuc[0] if sonuc else 0

    def urun_hucre_guncelle(self, urun_id, sutun_adi, yeni_deger):
        """Bir ürünün tek bir hücresini güvenli bir şekilde günceller."""
        # Sadece izin verilen sütunların güncellendiğinden emin ol
        izin_verilen_sutunlar = ['kategori', 'fiyat']
        if sutun_adi not in izin_verilen_sutunlar:
            return False, "Geçersiz güncelleme alanı."

        try:
            # f-string burada güvenli çünkü sutun_adi'ni beyaz listeyle kontrol ettik.
            self.cursor.execute(f"UPDATE urunler SET {sutun_adi} = ? WHERE id = ?",
                                (yeni_deger, urun_id))
            self.baglanti.commit()
            return True, "Güncellendi."
        except Exception as e:
            return False, f"Veritabanı hatası: {e}"

    def genel_bakis_getir(self):
        try:
            urun_cesidi = self.cursor.execute("SELECT COUNT(id) FROM urunler").fetchone()[0]
            toplam_stok_degeri = self.cursor.execute("SELECT SUM(fiyat * miktar) FROM urunler").fetchone()[0]
            dusuk_stok_sayisi = \
                self.cursor.execute("SELECT COUNT(id) FROM urunler WHERE miktar <= min_stok").fetchone()[0]
            return {
                "urun_cesidi": urun_cesidi or 0,
                "toplam_deger": toplam_stok_degeri or 0.0,
                "dusuk_stok": dusuk_stok_sayisi or 0
            }
        except Exception as e:
            print(f"!!! Genel bakış hatası (genel_bakis_getir): {e}")
            return {"urun_cesidi": 0, "toplam_deger": 0.0, "dusuk_stok": 0}

    def kullanici_sayisi_getir(self):
        self.cursor.execute("SELECT COUNT(*) FROM kullanicilar")
        return self.cursor.fetchone()[0]

    def sifre_hashle(self, s):
        return hashlib.sha256(s.encode()).hexdigest()

    def kullanici_ekle(self, k_adi, s):
        try:
            self.cursor.execute("INSERT INTO kullanicilar (kullanici_adi, sifre_hash) VALUES (?, ?)",
                                (k_adi, self.sifre_hashle(s)))
            self.baglanti.commit()
            return True, "Kullanıcı oluşturuldu."
        except sqlite3.IntegrityError:
            return False, "Bu kullanıcı adı zaten alınmış."

    def kullanici_dogrula(self, k_adi, s):
        self.cursor.execute("SELECT * FROM kullanicilar WHERE kullanici_adi = ? AND sifre_hash = ?",
                            (k_adi, self.sifre_hashle(s)))
        return self.cursor.fetchone() is not None

    def kullanici_bilgilerini_guncelle(self, e_kadi, y_kadi, y_sifre):
        try:
            self.cursor.execute("UPDATE kullanicilar SET kullanici_adi = ?, sifre_hash = ? WHERE kullanici_adi = ?",
                                (y_kadi, self.sifre_hashle(y_sifre), e_kadi))
            self.baglanti.commit()
            return True, "Bilgiler güncellendi."
        except sqlite3.IntegrityError:
            return False, "Yeni kullanıcı adı başkası tarafından kullanılıyor."


# Product Edit Dialog
class UrunDuzenlemeDialog(QDialog):
    def __init__(self, urun_detaylari, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ürün Bilgilerini Düzenle")
        self.form_layout = QFormLayout(self)

        self.urun_id = urun_detaylari[0]
        self.ad_input = QLineEdit(urun_detaylari[1])
        self.kategori_input = QLineEdit(urun_detaylari[2])
        self.fiyat_input = QLineEdit(str(urun_detaylari[3]))
        self.min_stok_input = QLineEdit(str(urun_detaylari[5]))

        self.form_layout.addRow("Ürün ID:", QLabel(str(self.urun_id)))
        self.form_layout.addRow("Ürün Adı:", self.ad_input)
        self.form_layout.addRow("Kategori:", self.kategori_input)
        self.form_layout.addRow("Fiyat (₺):", self.fiyat_input)
        self.form_layout.addRow("Min. Stok:", self.min_stok_input)

        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.form_layout.addWidget(self.buttonBox)

    def get_data(self):
        try:
            fiyat = float(self.fiyat_input.text().replace(',', '.'))
            min_stok = int(self.min_stok_input.text())
            ad = self.ad_input.text()
            kategori = self.kategori_input.text()
            return (self.urun_id, ad, kategori, fiyat, min_stok)
        except ValueError:
            QMessageBox.warning(self, "Hata", "Fiyat ve Min. Stok sayısal değer olmalıdır.")
            return None


# Advanced Filter Dialog
class FiltreDialog(QDialog):
    def __init__(self, mevcut_filtreler=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gelişmiş Filtrele")
        self.setMinimumWidth(400)

        layout = QFormLayout(self)
        layout.setSpacing(15)

        if mevcut_filtreler is None:
            mevcut_filtreler = {}

        # Fiyat Aralığı
        self.min_fiyat = QLineEdit(str(mevcut_filtreler.get("min_fiyat", "")))
        self.min_fiyat.setPlaceholderText("Minimum")
        self.max_fiyat = QLineEdit(str(mevcut_filtreler.get("max_fiyat", "")))
        self.max_fiyat.setPlaceholderText("Maksimum")
        fiyat_layout = QHBoxLayout()
        fiyat_layout.addWidget(self.min_fiyat)
        fiyat_layout.addWidget(QLabel("-"))
        fiyat_layout.addWidget(self.max_fiyat)
        layout.addRow("Fiyat Aralığı (₺):", fiyat_layout)

        # Stok Aralığı
        self.min_stok = QLineEdit(str(mevcut_filtreler.get("min_stok", "")))
        self.min_stok.setPlaceholderText("Minimum")
        self.max_stok = QLineEdit(str(mevcut_filtreler.get("max_stok", "")))
        self.max_stok.setPlaceholderText("Maksimum")
        stok_layout = QHBoxLayout()
        stok_layout.addWidget(self.min_stok)
        stok_layout.addWidget(QLabel("-"))
        stok_layout.addWidget(self.max_stok)
        layout.addRow("Stok Aralığı:", stok_layout)

        # Düşük Stok Seçeneği
        self.sadece_dusuk_stok = QCheckBox("Sadece düşük stoktakileri göster")
        self.sadece_dusuk_stok.setChecked(mevcut_filtreler.get("dusuk_stok_only", False))
        layout.addRow("", self.sadece_dusuk_stok)

        # Butonlar
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Reset)
        self.buttonBox.button(QDialogButtonBox.Ok).setText("Filtrele")
        self.buttonBox.button(QDialogButtonBox.Cancel).setText("İptal")
        self.buttonBox.button(QDialogButtonBox.Reset).setText("Filtreyi Temizle")

        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.buttonBox.button(QDialogButtonBox.Reset).clicked.connect(self.filtreyi_temizle_ve_kapat)

        layout.addWidget(self.buttonBox)

        self.reset_requested = False

    def filtreyi_temizle_ve_kapat(self):
        # Filtreleri temizle ve diyaloğun reset sinyaliyle kapanmasını sağla
        self.reset_requested = True
        self.accept()  # Reset'i de 'accept' gibi ele alacağız, ana pencere kontrol edecek

    def get_filtreler(self):
        """Diyalogdaki girdileri sayısal verilere dönüştürür."""

        def to_float(widget):
            try:
                return float(widget.text().replace(',', '.'))
            except ValueError:
                return None

        def to_int(widget):
            try:
                return int(widget.text())
            except ValueError:
                return None

        return {
            "min_fiyat": to_float(self.min_fiyat),
            "max_fiyat": to_float(self.max_fiyat),
            "min_stok": to_int(self.min_stok),
            "max_stok": to_int(self.max_stok),
            "dusuk_stok_only": self.sadece_dusuk_stok.isChecked()
        }


# Main Application Widget
class StokUygulamasiWidget(QWidget):
    def __init__(self, veritabani_yoneticisi, status_bar):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.status_bar = status_bar
        self.style_ikon = QApplication.style()
        self.guncel_filtreler = {}  # Gelişmiş filtre ayarlarını burada tutacağız
        self.arayuz_olustur()
        self.stogu_guncelle_arayuz()

    def create_metric_card(self, parent_layout, icon, title, value, unit):
        """Create a modern metric card for the dashboard"""
        card_frame = QFrame()
        card_frame.setObjectName("metricCard")
        card_layout = QVBoxLayout(card_frame)
        card_layout.setAlignment(Qt.AlignCenter)
        card_layout.setSpacing(8)

        # Icon and title
        icon_label = QLabel(icon)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("font-size: 24pt; margin-bottom: 4px;")

        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 10pt; font-weight: 500; opacity: 0.9;")

        # Value and unit
        value_layout = QHBoxLayout()
        value_layout.setAlignment(Qt.AlignCenter)

        value_label = QLabel(value)
        value_label.setObjectName("metricValue")
        value_label.setAlignment(Qt.AlignCenter)
        value_label.setStyleSheet("font-size: 18pt; font-weight: 700; margin: 0;")

        unit_label = QLabel(unit)
        unit_label.setAlignment(Qt.AlignCenter)
        unit_label.setStyleSheet("font-size: 9pt; font-weight: 500; opacity: 0.8; margin-left: 4px;")

        value_layout.addWidget(value_label)
        value_layout.addWidget(unit_label)

        card_layout.addWidget(icon_label)
        card_layout.addWidget(title_label)
        card_layout.addLayout(value_layout)

        # Store references for updating
        if not hasattr(self, 'metric_cards'):
            self.metric_cards = {}

        if title == "Ürün Çeşidi":
            self.metric_cards['urun_cesidi'] = value_label
        elif title == "Toplam Değer":
            self.metric_cards['toplam_deger'] = value_label
        elif title == "Düşük Stok":
            self.metric_cards['dusuk_stok'] = value_label
        elif title == "Ortalama Fiyat":
            self.metric_cards['ortalama_fiyat'] = value_label

        parent_layout.addWidget(card_frame, 1)

    def arayuz_olustur(self):
        ana_duzen = QVBoxLayout(self)
        ana_duzen.setContentsMargins(15, 15, 15, 15)
        ana_duzen.setSpacing(10)

        # Modern Dashboard Panel
        dashboard_frame = QFrame()
        dashboard_frame.setObjectName("dashboardFrame")
        dashboard_frame.setFixedHeight(120)
        dashboard_layout = QHBoxLayout(dashboard_frame)
        dashboard_layout.setSpacing(20)

        # Create individual metric cards
        self.create_metric_card(dashboard_layout, "📦", "Ürün Çeşidi", "0", "adet")
        self.create_metric_card(dashboard_layout, "💰", "Toplam Değer", "0", "₺")
        self.create_metric_card(dashboard_layout, "⚠️", "Düşük Stok", "0", "ürün")
        self.create_metric_card(dashboard_layout, "📈", "Ortalama Fiyat", "0", "₺")

        ana_duzen.addWidget(dashboard_frame)

        # Modern Search and Action Bar
        ust_duzen = QHBoxLayout()
        ust_duzen.setSpacing(12)

        # Search container
        search_container = QFrame()
        search_container.setObjectName("searchContainer")
        search_layout = QHBoxLayout(search_container)
        search_layout.setContentsMargins(12, 8, 12, 8)

        self.arama_input = QLineEdit()
        self.arama_input.setPlaceholderText("🔍 Ürün adı, kategori veya fiyat ile ara...")
        self.arama_input.setObjectName("searchInput")
        # Filtreleme fonksiyonunu lambda ile bağla
        self.arama_input.textChanged.connect(lambda: self.tabloyu_filtrele())
        search_layout.addWidget(self.arama_input)

        # Clear search button
        clear_btn = QPushButton("✕")
        clear_btn.setObjectName("clearSearchBtn")
        clear_btn.setFixedSize(30, 30)
        clear_btn.setToolTip("Aramayı temizle")
        clear_btn.clicked.connect(self.arama_input.clear)
        search_layout.addWidget(clear_btn)

        ust_duzen.addWidget(search_container, 1)

        # Action buttons container
        action_container = QHBoxLayout()
        action_container.setSpacing(8)

        # Filter button
        self.filter_btn = QPushButton("🔽 Filtrele")
        self.filter_btn.setObjectName("filterBtn")
        self.filter_btn.clicked.connect(self.show_filter_dialog)
        action_container.addWidget(self.filter_btn)

        # Add product button
        self.yeni_urun_goster_btn = QPushButton("➕ Yeni Ürün")
        self.yeni_urun_goster_btn.setObjectName("yeniUrunBtn")
        self.yeni_urun_goster_btn.clicked.connect(self.yeni_urun_formu_goster_gizle)
        action_container.addWidget(self.yeni_urun_goster_btn)

        ust_duzen.addLayout(action_container)
        ana_duzen.addLayout(ust_duzen)

        # Product Addition Form
        self.ekleme_formu_frame = QFrame()
        self.ekleme_formu_frame.setObjectName("eklemeFormu")
        self.ekleme_formu_frame.setFrameShape(QFrame.StyledPanel)
        ekleme_duzen = QFormLayout(self.ekleme_formu_frame)
        ekleme_duzen.setContentsMargins(10, 10, 10, 10)
        ekleme_duzen.setSpacing(8)

        self.yeni_urun_input = QLineEdit()
        self.yeni_kategori_input = QLineEdit()
        self.yeni_fiyat_input = QLineEdit()
        self.yeni_miktar_input = QLineEdit()
        self.yeni_min_stok_input = QLineEdit("10")

        self.onayla_ekle_btn = QPushButton("✅ Onayla")
        self.onayla_ekle_btn.setDefault(True)

        ekleme_duzen.addRow("Ürün Adı:", self.yeni_urun_input)
        ekleme_duzen.addRow("Kategori:", self.yeni_kategori_input)
        ekleme_duzen.addRow("Fiyat (₺):", self.yeni_fiyat_input)
        ekleme_duzen.addRow("Miktar:", self.yeni_miktar_input)
        ekleme_duzen.addRow("Min. Stok:", self.yeni_min_stok_input)
        ekleme_duzen.addRow(self.onayla_ekle_btn)

        self.ekleme_formu_frame.hide()
        ana_duzen.addWidget(self.ekleme_formu_frame)
        self.onayla_ekle_btn.clicked.connect(self.yeni_urun_ekle)

        # Main Stock Table
        self.stok_tablosu = QTableWidget()
        self.stok_tablosu.setColumnCount(7)
        self.stok_tablosu.setHorizontalHeaderLabels(
            ["ID", "Ürün Adı", "Kategori", "Fiyat", "Miktar", "Min. Stok", "İşlemler"])

        # Column width settings
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)

        self.stok_tablosu.setAlternatingRowColors(True)
        self.stok_tablosu.setShowGrid(False)
        self.stok_tablosu.verticalHeader().setVisible(False)
        self.stok_tablosu.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.stok_tablosu.setSelectionMode(QAbstractItemView.SingleSelection)

        # Hızlı Düzenleme için değiştirildi
        self.stok_tablosu.setEditTriggers(QAbstractItemView.DoubleClicked)
        # Hızlı Düzenleme için sinyal eklendi
        self.stok_tablosu.itemChanged.connect(self.hucre_degisikligini_kaydet)

        # Sıralama Özellikleri Eklendi
        self.stok_tablosu.setSortingEnabled(True)
        self.stok_tablosu.sortByColumn(1, Qt.AscendingOrder)  # Varsayılan olarak Ada göre sırala

        ana_duzen.addWidget(self.stok_tablosu, 1)

    def guncelle_dashboard(self):
        """Update dashboard metrics"""
        veri = self.veritabani.genel_bakis_getir()

        # Calculate average price
        ortalama_fiyat = veri['toplam_deger'] / veri['urun_cesidi'] if veri['urun_cesidi'] > 0 else 0

        # Update metric cards
        if hasattr(self, 'metric_cards'):
            self.metric_cards['urun_cesidi'].setText(str(veri['urun_cesidi']))
            self.metric_cards['toplam_deger'].setText(f"{veri['toplam_deger']:,.0f}")
            self.metric_cards['dusuk_stok'].setText(str(veri['dusuk_stok']))
            self.metric_cards['ortalama_fiyat'].setText(f"{ortalama_fiyat:,.0f}")

            # Color coding for low stock warning
            if veri['dusuk_stok'] > 0:
                self.metric_cards['dusuk_stok'].setStyleSheet(
                    "font-size: 18pt; font-weight: 700; margin: 0; color: #ef4444;")
            else:
                self.metric_cards['dusuk_stok'].setStyleSheet(
                    "font-size: 18pt; font-weight: 700; margin: 0; color: white;")

    def stogu_guncelle_arayuz(self):
        """Update the interface with current data"""
        self.guncelle_dashboard()

        # Hızlı Düzenleme: Sinyali geçici olarak durdur
        # Tabloyu doldururken 'hucre_degisikligini_kaydet' fonksiyonunun tetiklenmesini engelle
        try:
            self.stok_tablosu.itemChanged.disconnect(self.hucre_degisikligini_kaydet)
        except TypeError:
            pass  # Sinyal henüz bağlanmamışsa (ilk çalıştırma)

        # Tablonun sıralama yaparken sinyal göndermesini geçici olarak engelle
        self.stok_tablosu.setSortingEnabled(False)

        self.stok_tablosu.setRowCount(0)
        urun_listesi = self.veritabani.urunleri_getir()
        self.stok_tablosu.setRowCount(len(urun_listesi))

        urun_fontu = QFont("Segoe UI Semibold", 11)
        normal_font = QFont("Segoe UI", 10)
        id_fontu = QFont("Segoe UI", 9)

        id_renk = QColor("#A0A0A0")
        dusuk_stok_renk = QColor(255, 100, 100, 40)

        for satir, (id_val, ad, kategori, fiyat, miktar, min_stok) in enumerate(urun_listesi):
            # ID cell
            id_item = NumericTableWidgetItem(str(id_val), id_val)
            id_item.setFont(id_fontu)
            id_item.setForeground(id_renk)
            id_item.setTextAlignment(Qt.AlignCenter)
            id_item.setData(Qt.UserRole, id_val)  # Menü için ID'yi sakla

            # Other cells
            ad_item = QTableWidgetItem(ad)  # Metin sıralaması normal
            ad_item.setFont(urun_fontu)

            kategori_item = QTableWidgetItem(kategori)  # Metin sıralaması normal
            kategori_item.setFont(normal_font)

            fiyat_item = NumericTableWidgetItem(f"{fiyat:.2f} ₺", fiyat)
            fiyat_item.setFont(normal_font)
            fiyat_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            miktar_item = NumericTableWidgetItem(str(miktar), miktar)
            miktar_item.setFont(normal_font)
            miktar_item.setTextAlignment(Qt.AlignCenter)

            min_stok_item = NumericTableWidgetItem(str(min_stok), min_stok)
            min_stok_item.setFont(normal_font)
            min_stok_item.setTextAlignment(Qt.AlignCenter)

            # --- Hızlı Düzenleme: Hücre İzinleri ---
            # İstenmeyen hücrelerin düzenlenmesini engelle
            id_item.setFlags(id_item.flags() & ~Qt.ItemIsEditable)
            ad_item.setFlags(ad_item.flags() & ~Qt.ItemIsEditable)
            miktar_item.setFlags(miktar_item.flags() & ~Qt.ItemIsEditable)
            min_stok_item.setFlags(min_stok_item.flags() & ~Qt.ItemIsEditable)
            # Kategori (kategori_item) ve Fiyat (fiyat_item) varsayılan olarak düzenlenebilir kalır.

            self.stok_tablosu.setItem(satir, 0, id_item)
            self.stok_tablosu.setItem(satir, 1, ad_item)
            self.stok_tablosu.setItem(satir, 2, kategori_item)
            self.stok_tablosu.setItem(satir, 3, fiyat_item)
            self.stok_tablosu.setItem(satir, 4, miktar_item)
            self.stok_tablosu.setItem(satir, 5, min_stok_item)

            # Low stock warning - Modern styling
            if miktar <= min_stok:
                for col in range(self.stok_tablosu.columnCount() - 1):
                    item = self.stok_tablosu.item(satir, col)
                    item.setBackground(QColor(254, 242, 242))
                    item.setForeground(QColor(220, 38, 38))
                    if col == 4:  # Quantity column
                        item.setText(f"⚠️ {miktar}")

            # Modern Action Button
            menu_btn = QPushButton("⋮")
            menu_btn.setFixedSize(36, 32)
            menu_btn.setObjectName("menuButton")
            menu_btn.setToolTip("İşlemler")

            # --- SIRALAMA BUG FİX ---
            # Butonun tıklanma olayını, satır indeksine değil, doğrudan 'id_val' ve 'ad'a bağla.
            menu_btn.clicked.connect(
                lambda ch, urun_id=id_val, ad=ad, btn=menu_btn: self.guncelle_menusu_goster(urun_id, ad, btn)
            )
            # -------------------------

            self.stok_tablosu.setCellWidget(satir, 6, menu_btn)

        self.stok_tablosu.resizeRowsToContents()
        self.stok_tablosu.setSortingEnabled(True)  # Sıralamayı tekrar aç
        self.tabloyu_filtrele()  # Arayüz güncellendikten sonra filtreleri uygula

        # Hızlı Düzenleme: Sinyali yeniden bağla
        self.stok_tablosu.itemChanged.connect(self.hucre_degisikligini_kaydet)

    def guncelle_menusu_goster(self, urun_id, urun_adi_gosterim, buton):
        """Show context menu for product actions"""
        # --- SIRALAMA BUG FİX ---
        # Artık satır indeksine göre arama yapmaya gerek yok,
        # 'urun_id' ve 'urun_adi_gosterim' doğrudan parametre olarak geliyor.

        menu = QMenu(self)

        # Modern menu actions with emojis
        duzenle = QAction("✏️ Bilgileri Düzenle (Detay)", self)
        duzenle.triggered.connect(lambda: self.urun_duzenle(urun_id))
        menu.addAction(duzenle)

        menu.addSeparator()

        artir = QAction("📈 Stok Artır", self)
        artir.triggered.connect(lambda: self.miktar_girdi_goster(urun_id, urun_adi_gosterim, 'artır'))
        menu.addAction(artir)

        eksilt = QAction("📉 Stok Eksilt", self)
        eksilt.triggered.connect(lambda: self.miktar_girdi_goster(urun_id, urun_adi_gosterim, 'eksilt'))
        menu.addAction(eksilt)

        menu.addSeparator()

        sil = QAction("🗑️ Ürünü Sil", self)
        sil.setObjectName("dangerAction")
        sil.triggered.connect(lambda: self.urun_sil(urun_id, urun_adi_gosterim))
        menu.addAction(sil)

        menu.exec_(buton.mapToGlobal(buton.rect().bottomLeft()))

    def urun_duzenle(self, urun_id):
        """Edit product information"""
        urun_detaylari = self.veritabani.urun_detay_getir(urun_id)
        if not urun_detaylari:
            QMessageBox.critical(self, "Hata", "Ürün detayı bulunamadı.")
            return

        dialog = UrunDuzenlemeDialog(urun_detaylari, self)
        if dialog.exec_() == QDialog.Accepted:
            veri = dialog.get_data()
            if veri:
                id_val, yeni_ad, yeni_kategori, yeni_fiyat, yeni_min_stok = veri
                self.veritabani.urun_detay_guncelle(id_val, yeni_ad, yeni_kategori, yeni_fiyat, yeni_min_stok)
                self.stogu_guncelle_arayuz()
                self.status_bar.showMessage(f"'{yeni_ad}' (ID: {id_val}) güncellendi.", 3000)

    def yeni_urun_ekle(self):
        """Add new product"""
        ad = self.yeni_urun_input.text().strip().capitalize()
        kat = self.yeni_kategori_input.text().strip()
        f_str = self.yeni_fiyat_input.text().strip()
        m_str = self.yeni_miktar_input.text().strip()
        min_s_str = self.yeni_min_stok_input.text().strip()

        if not all([ad, f_str, m_str, min_s_str]):
            QMessageBox.warning(self, "Uyarı", "Ürün Adı, Fiyat, Miktar ve Min. Stok zorunludur.")
            return

        try:
            f = float(f_str.replace(',', '.'))
            m = int(m_str)
            min_s = int(min_s_str)
            if not (m > 0 and f >= 0 and min_s >= 0):
                raise ValueError("Değerler pozitif olmalı")
        except(ValueError):
            QMessageBox.warning(self, "Hata", "Fiyat, Miktar ve Min. Stok pozitif sayı olmalıdır.")
            return

        self.veritabani.urun_ekle(ad, kat, f, m, min_s)
        self.stogu_guncelle_arayuz()
        self.yeni_urun_formu_goster_gizle()
        self.status_bar.showMessage(f"'{ad}' eklendi/güncellendi.", 3000)

        # Clear fields
        for item in [self.yeni_urun_input, self.yeni_kategori_input, self.yeni_fiyat_input,
                     self.yeni_miktar_input, self.yeni_min_stok_input]:
            item.clear()
        self.yeni_min_stok_input.setText("10")

    def show_filter_dialog(self):
        """Gelişmiş filtre diyalogunu göster"""
        dialog = FiltreDialog(self.guncel_filtreler, self)

        if dialog.exec_() == QDialog.Accepted:
            if dialog.reset_requested:
                # Kullanıcı "Filtreyi Temizle" butonuna bastı
                self.guncel_filtreler = {}
                self.status_bar.showMessage("Filtreler temizlendi.", 3000)
                self.filter_btn.setText("🔽 Filtrele")
                self.filter_btn.setStyleSheet("")  # Stili sıfırla
            else:
                # Kullanıcı "Filtrele" butonuna bastı
                self.guncel_filtreler = dialog.get_filtreler()
                self.status_bar.showMessage("Gelişmiş filtreler uygulandı.", 3000)
                # Butonun görünümünü değiştirerek filtrenin aktif olduğunu belirt
                self.filter_btn.setText("✔️ Filtreleniyor")
                self.filter_btn.setStyleSheet(
                    "background: #fef9c3; color: #713f12; border-color: #fde047;"
                )

        # Hem 'Filtrele' hem de 'Temizle' durumunda filtrelemeyi yeniden çalıştır
        self.tabloyu_filtrele()

    def tabloyu_filtrele(self, metin=None):
        """
        Gelişmiş filtre fonksiyonu. Hem arama çubuğunu
        hem de gelişmiş filtre ayarlarını kontrol eder.
        """
        metin_lower = self.arama_input.text().lower()

        # Gelişmiş filtre ayarlarını al
        min_f = self.guncel_filtreler.get("min_fiyat")
        max_f = self.guncel_filtreler.get("max_fiyat")
        min_s = self.guncel_filtreler.get("min_stok")
        max_s = self.guncel_filtreler.get("max_stok")
        dusuk_stok_only = self.guncel_filtreler.get("dusuk_stok_only", False)

        for i in range(self.stok_tablosu.rowCount()):
            # 1. Metin Arama Kontrolü
            urun_adi = self.stok_tablosu.item(i, 1).text().lower()
            kategori = self.stok_tablosu.item(i, 2).text().lower()
            fiyat_text = self.stok_tablosu.item(i, 3).text().lower()  # Formatlı metni ara

            text_match = (metin_lower in urun_adi or
                          metin_lower in kategori or
                          metin_lower in fiyat_text)

            # 2. Gelişmiş Filtre Kontrolü
            # NumericTableWidgetItem'den ham sayısal verileri al
            try:
                fiyat = self.stok_tablosu.item(i, 3).sort_key
                miktar = self.stok_tablosu.item(i, 4).sort_key
                min_stok_val = self.stok_tablosu.item(i, 5).sort_key
            except Exception:
                # Bir hata olursa (örn: öğe düzgün oluşturulmadıysa) bu satırı atla
                self.stok_tablosu.setRowHidden(i, True)
                continue

            # Fiyat aralığı kontrolü
            fiyat_match = True
            if min_f is not None and fiyat < min_f:
                fiyat_match = False
            if max_f is not None and fiyat > max_f:
                fiyat_match = False

            # Stok aralığı kontrolü
            stok_match = True
            if min_s is not None and miktar < min_s:
                stok_match = False
            if max_s is not None and miktar > max_s:
                stok_match = False

            # Düşük stok kontrolü
            dusuk_stok_match = True
            if dusuk_stok_only and miktar > min_stok_val:
                dusuk_stok_match = False

            # --- SONUÇ ---
            # Tüm koşullar sağlanıyorsa satırı göster
            is_visible = (text_match and fiyat_match and stok_match and dusuk_stok_match)
            self.stok_tablosu.setRowHidden(i, not is_visible)

    def yeni_urun_formu_goster_gizle(self):
        """Toggle product addition form visibility"""
        if self.ekleme_formu_frame.isVisible():
            self.ekleme_formu_frame.hide()
        else:
            self.ekleme_formu_frame.show()
            self.yeni_urun_input.setFocus()

    def miktar_girdi_goster(self, urun_id, urun_adi, mod):
        """Show quantity input dialog"""
        fiil = "Artır" if mod == 'artır' else "Eksilt"
        mevcut = self.veritabani.mevcut_miktar_getir(urun_id)

        m, ok = QInputDialog.getInt(self, f"Stok {fiil} - [{urun_adi}]",
                                    f"Mevcut Miktar: {mevcut}\nLütfen miktarı girin:", 1, 1, 999999)
        if ok:
            self.stok_miktari_degistir(urun_id, urun_adi, m if mod == 'artır' else -m)

    def stok_miktari_degistir(self, urun_id, urun_adi, m_farki):
        """Change product quantity"""
        mevcut = self.veritabani.mevcut_miktar_getir(urun_id)
        if mevcut + m_farki < 0:
            QMessageBox.warning(self, "Uyarı", f"Stok eksiye düşemez. (Mevcut: {mevcut})")
            return

        self.veritabani.urun_miktar_guncelle(urun_id, m_farki)
        self.stogu_guncelle_arayuz()
        self.status_bar.showMessage(f"'{urun_adi}' stoğu güncellendi.", 3000)

    def urun_sil(self, urun_id, urun_adi):
        """Delete product"""
        onay = QMessageBox.question(self, "Silmeyi Onayla",
                                    f"<b>{urun_adi}</b> (ID: {urun_id}) ürününü silmek istediğinizden emin misiniz?",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if onay == QMessageBox.Yes:
            self.veritabani.urun_sil(urun_id)
            self.stogu_guncelle_arayuz()
            self.status_bar.showMessage(f"'{urun_adi}' silindi.", 3000)

    def verileri_disa_aktar(self):
        """Export data to CSV"""
        kayit_yolu, _ = QFileDialog.getSaveFileName(self, "CSV Olarak Dışa Aktar",
                                                    "stok_raporu.csv",
                                                    "CSV Dosyaları (*.csv)")

        if not kayit_yolu:
            return

        try:
            urunler = self.veritabani.urunleri_getir()
            basliklar = ["ID", "Ad", "Kategori", "Fiyat", "Miktar", "Min. Stok"]

            with open(kayit_yolu, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(basliklar)
                writer.writerows(urunler)

            self.status_bar.showMessage(f"Veriler '{kayit_yolu}' dosyasına aktarıldı.", 5000)
            QMessageBox.information(self, "Başarılı", "Tüm veriler CSV dosyasına başarıyla aktarıldı.")

        except Exception as e:
            QMessageBox.critical(self, "Dışa Aktarma Hatası", f"Dosya yazılırken bir hata oluştu:\n{e}")

    def hucre_degisikligini_kaydet(self, item):
        """Tablodan yapılan çift tıklama değişikliklerini veritabanına kaydeder."""
        satir = item.row()
        sutun = item.column()

        # Değişiklik yapılan satırın ID'sini al
        try:
            id_item = self.stok_tablosu.item(satir, 0)
            if id_item is None:
                # Hücre (belki filtrelenmiş) geçerli bir satıra ait değilse
                return
            urun_id = id_item.data(Qt.UserRole)
            urun_adi = self.stok_tablosu.item(satir, 1).text()
        except Exception as e:
            # ID'yi alamazsak hiçbir şey yapamayız
            print(f"Hata: ID alınamadı (satir {satir}): {e}")
            return

        yeni_deger_str = item.text()

        # Hangi Sütun Değişti?
        if sutun == 2:  # Kategori
            basari, mesaj = self.veritabani.urun_hucre_guncelle(urun_id, "kategori", yeni_deger_str)
            if basari:
                self.status_bar.showMessage(f"'{urun_adi}' kategorisi '{yeni_deger_str}' olarak güncellendi.", 3000)
            else:
                QMessageBox.warning(self, "Hata", mesaj)
                # Başarısız olursa değişikliği geri al (tabloyu yeniden yükleyerek)
                self.stogu_guncelle_arayuz()

        elif sutun == 3:  # Fiyat
            try:
                # '150,50 ₺' veya '150.5' gibi girdileri temizle
                temiz_deger = yeni_deger_str.replace('₺', '').replace(',', '.').strip()
                yeni_fiyat = float(temiz_deger)

                if yeni_fiyat < 0:
                    raise ValueError("Fiyat negatif olamaz")

                # Veritabanını güncelle
                basari, mesaj = self.veritabani.urun_hucre_guncelle(urun_id, "fiyat", yeni_fiyat)

                if basari:
                    self.status_bar.showMessage(f"'{urun_adi}' fiyatı {yeni_fiyat:.2f} ₺ olarak güncellendi.", 3000)
                    # Dashboard'u yeni toplam değer için güncelle
                    self.guncelle_dashboard()

                    # Hücreyi yeniden formatla (Sinyal yönetimi gerekli)
                    self.stok_tablosu.itemChanged.disconnect(self.hucre_degisikligini_kaydet)

                    yeni_fiyat_item = NumericTableWidgetItem(f"{yeni_fiyat:.2f} ₺", yeni_fiyat)
                    yeni_fiyat_item.setFont(QFont("Segoe UI", 10))
                    yeni_fiyat_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    # Sadece Fiyat hücresinin bayraklarını ayarla
                    yeni_fiyat_item.setFlags(yeni_fiyat_item.flags() | Qt.ItemIsEditable)
                    self.stok_tablosu.setItem(satir, sutun, yeni_fiyat_item)

                    self.stok_tablosu.itemChanged.connect(self.hucre_degisikligini_kaydet)
                else:
                    QMessageBox.warning(self, "Hata", mesaj)
                    self.stogu_guncelle_arayuz()  # Değişikliği geri al

            except ValueError:
                # Eğer kullanıcı 'abc' gibi geçersiz bir fiyat girerse
                QMessageBox.warning(self, "Geçersiz Değer", "Lütfen fiyat için geçerli bir pozitif sayı girin.")
                self.stogu_guncelle_arayuz()  # Değişikliği geri al


# Main Window
class AnaPencere(QMainWindow):
    bilgi_degistir_iste = pyqtSignal()

    def __init__(self, kullanici_adi, veritabani_yoneticisi):
        super().__init__()
        self.setWindowTitle(f"📦 Stok Yönetim Sistemi - {kullanici_adi}")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1200, 700)

        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Hoş geldiniz!", 3000)

        self.merkezi_widget = StokUygulamasiWidget(veritabani_yoneticisi, self.status_bar)
        self.setCentralWidget(self.merkezi_widget)

        self.style_ikon = QApplication.style()

        # Menu Bar
        self.menu_bar = self.menuBar()
        self.menu_bar.setObjectName("menuBar")

        # File Menu
        dosya_menu = self.menu_bar.addMenu("📁 Dosya")
        disa_aktar_action = QAction("💾 Verileri CSV Olarak Dışa Aktar...", self)
        disa_aktar_action.triggered.connect(self.merkezi_widget.verileri_disa_aktar)
        dosya_menu.addAction(disa_aktar_action)

        cikis_action = QAction("🚪 Çıkış Yap", self)
        cikis_action.triggered.connect(self.close)
        dosya_menu.addAction(cikis_action)

        # Settings Menu
        ayarlar_menu = self.menu_bar.addMenu("⚙️ Ayarlar")
        kullanici_degistir_action = QAction("👤 Kullanıcı Bilgilerini Değiştir...", self)
        kullanici_degistir_action.triggered.connect(self.bilgi_degistir_iste.emit)
        ayarlar_menu.addAction(kullanici_degistir_action)


# Login and Setup Windows
class IlkKurulumPenceresi(QWidget):
    kurulum_tamamlandi = pyqtSignal(str)

    def __init__(self, veritabani_yoneticisi):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.setWindowTitle("İlk Kurulum")
        self.setGeometry(400, 400, 400, 250)
        duzen = QVBoxLayout(self)
        duzen.addWidget(QLabel("<h2>Yönetici Hesabı Oluştur</h2>"))
        self.k_adi = QLineEdit()
        self.k_adi.setPlaceholderText("Kullanıcı Adı")
        self.sifre = QLineEdit()
        self.sifre.setPlaceholderText("Şifre")
        self.sifre.setEchoMode(QLineEdit.Password)
        self.sifre_t = QLineEdit()
        self.sifre_t.setPlaceholderText("Şifre Tekrar")
        self.sifre_t.setEchoMode(QLineEdit.Password)
        btn = QPushButton("✅ Hesabı Oluştur")
        btn.setDefault(True)
        btn.clicked.connect(self.hesap_olustur)
        duzen.addWidget(self.k_adi)
        duzen.addWidget(self.sifre)
        duzen.addWidget(self.sifre_t)
        duzen.addWidget(btn)

    def hesap_olustur(self):
        k, s, st = self.k_adi.text().strip(), self.sifre.text(), self.sifre_t.text()
        if not k or not s:
            QMessageBox.warning(self, "Hata", "Alanlar boş bırakılamaz.")
            return
        if s != st:
            QMessageBox.warning(self, "Hata", "Şifreler eşleşmiyor.")
            return
        b, m = self.veritabani.kullanici_ekle(k, s)
        if b:
            self.kurulum_tamamlandi.emit(k)
            self.close()
        else:
            QMessageBox.critical(self, "Hata", m)


class KullaniciDegistirPenceresi(QWidget):
    degisiklik_yapildi = pyqtSignal()

    def __init__(self, veritabani_yoneticisi):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.setWindowTitle("Bilgileri Değiştir")
        self.setGeometry(400, 300, 450, 400)
        duzen = QVBoxLayout(self)
        duzen.addWidget(QLabel("<h3>Mevcut Bilgileri Doğrula</h3>"))
        self.e_kadi = QLineEdit()
        self.e_kadi.setPlaceholderText("Mevcut Kullanıcı Adı")
        self.e_sifre = QLineEdit()
        self.e_sifre.setPlaceholderText("Mevcut Şifre")
        self.e_sifre.setEchoMode(QLineEdit.Password)
        duzen.addWidget(self.e_kadi)
        duzen.addWidget(self.e_sifre)
        duzen.addWidget(QLabel("<h3>Yeni Bilgileri Girin</h3>"))
        self.y_kadi = QLineEdit()
        self.y_kadi.setPlaceholderText("Yeni Kullanıcı Adı")
        self.y_sifre = QLineEdit()
        self.y_sifre.setPlaceholderText("Yeni Şifre")
        self.y_sifre.setEchoMode(QLineEdit.Password)
        self.y_sifre_t = QLineEdit()
        self.y_sifre_t.setPlaceholderText("Yeni Şifre Tekrar")
        self.y_sifre_t.setEchoMode(QLineEdit.Password)
        duzen.addWidget(self.y_kadi)
        duzen.addWidget(self.y_sifre)
        duzen.addWidget(self.y_sifre_t)
        btn = QPushButton("✅ Değişiklikleri Onayla")
        btn.setDefault(True)
        btn.clicked.connect(self.bilgileri_degistir)
        duzen.addStretch()
        duzen.addWidget(btn)

    def bilgileri_degistir(self):
        ek, es = self.e_kadi.text().strip(), self.e_sifre.text()
        yk, ys, yst = self.y_kadi.text().strip(), self.y_sifre.text(), self.y_sifre_t.text()
        if not all([ek, es, yk, ys]):
            QMessageBox.warning(self, "Hata", "Tüm alanlar doldurulmalıdır.")
            return
        if ys != yst:
            QMessageBox.warning(self, "Hata", "Yeni şifreler eşleşmiyor.")
            return
        if not self.veritabani.kullanici_dogrula(ek, es):
            QMessageBox.warning(self, "Hata", "Mevcut bilgiler yanlış.")
            return
        b, m = self.veritabani.kullanici_bilgilerini_guncelle(ek, yk, ys)
        if b:
            QMessageBox.information(self, "Başarılı", m)
            self.degisiklik_yapildi.emit()
            self.close()
        else:
            QMessageBox.warning(self, "Hata", m)


class GirisPenceresi(QWidget):
    login_basarili = pyqtSignal(str)
    degistirme_penceresi_iste = pyqtSignal()

    def __init__(self, veritabani_yoneticisi):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.setWindowTitle("Giriş")
        self.setGeometry(400, 400, 400, 250)
        duzen = QVBoxLayout(self)
        duzen.addWidget(QLabel("<h2>Giriş Yap</h2>"))
        self.k_adi = QLineEdit()
        self.k_adi.setPlaceholderText("Kullanıcı Adı")
        self.sifre = QLineEdit()
        self.sifre.setPlaceholderText("Şifre")
        self.sifre.setEchoMode(QLineEdit.Password)
        login_btn = QPushButton("🔑 Giriş Yap")
        login_btn.setDefault(True)
        login_btn.clicked.connect(self.login_kontrol)
        degistir_btn = QPushButton("Bilgileri Değiştir")
        degistir_btn.setObjectName("degistirBtn")
        degistir_btn.setFlat(True)
        degistir_btn.clicked.connect(self.degistirme_penceresi_iste.emit)
        duzen.addWidget(self.k_adi)
        duzen.addWidget(self.sifre)
        duzen.addWidget(login_btn)
        duzen.addWidget(degistir_btn, 0, Qt.AlignRight)
        self.sifre.returnPressed.connect(self.login_kontrol)

    def login_kontrol(self):
        k, s = self.k_adi.text(), self.sifre.text()
        if self.veritabani.kullanici_dogrula(k, s):
            self.login_basarili.emit(k)
            self.close()
        else:
            QMessageBox.warning(self, "Hata", "Kullanıcı adı veya şifre hatalı.")


# Main Controller
class AnaKontrolcu:
    def __init__(self):
        self.veritabani = VeritabaniYoneticisi()
        self.login_penceresi = None
        self.ana_pencere = None
        self.degistirme_penceresi = None
        self.mevcut_pencere = None

    def baslat(self):
        if self.veritabani.kullanici_sayisi_getir() == 0:
            self.mevcut_pencere = IlkKurulumPenceresi(self.veritabani)
            self.mevcut_pencere.kurulum_tamamlandi.connect(self.ana_pencereyi_goster)
        else:
            self.login_penceresi = GirisPenceresi(self.veritabani)
            self.login_penceresi.login_basarili.connect(self.ana_pencereyi_goster)
            self.login_penceresi.degistirme_penceresi_iste.connect(self.degistirme_penceresini_goster_login)
            self.mevcut_pencere = self.login_penceresi
        self.mevcut_pencere.show()

    def ana_pencereyi_goster(self, kullanici_adi):
        self.ana_pencere = AnaPencere(kullanici_adi, self.veritabani)
        self.ana_pencere.bilgi_degistir_iste.connect(self.degistirme_penceresini_goster_ana)
        self.ana_pencere.show()
        if self.login_penceresi:
            self.login_penceresi.close()
        if self.mevcut_pencere:
            self.mevcut_pencere.close()
        self.mevcut_pencere = None
        self.login_penceresi = None

    def degistirme_penceresini_goster_login(self):
        self.degistirme_penceresi = KullaniciDegistirPenceresi(self.veritabani)
        self.degistirme_penceresi.degisiklik_yapildi.connect(self.login_penceresine_don)
        self.degistirme_penceresi.show()
        if self.login_penceresi:
            self.login_penceresi.hide()

    def degistirme_penceresini_goster_ana(self):
        if self.degistirme_penceresi is None or not self.degistirme_penceresi.isVisible():
            self.degistirme_penceresi = KullaniciDegistirPenceresi(self.veritabani)
            self.degistirme_penceresi.show()
        self.degistirme_penceresi.activateWindow()

    def login_penceresine_don(self):
        if self.degistirme_penceresi:
            self.degistirme_penceresi.close()
        if self.login_penceresi:
            self.login_penceresi.show()


# Application Startup
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Modern Professional Theme
    app.setStyleSheet("""
        /* Modern Color Palette */
        * {
            font-family: 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
            font-size: 9pt;
        }

        /* Main Application */
        QMainWindow, QWidget, QDialog {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #f8fafc, stop:1 #f1f5f9);
            color: #1e293b;
            border: none;
        }

        /* Modern Header */
        QFrame#dashboardFrame {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #3b82f6, stop:1 #1d4ed8);
            border-radius: 12px;
            border: none;
            padding: 20px;
            margin: 10px;
        }

        QFrame#dashboardFrame QLabel {
            color: white;
            font-size: 11pt;
            font-weight: 500;
        }

        QFrame#dashboardFrame QLabel b {
            color: white;
            font-size: 14pt;
            font-weight: 700;
        }

        /* Modern Input Fields */
        QLineEdit {
            background-color: white;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            padding: 12px 16px;
            font-size: 10pt;
            color: #1e293b;
            selection-background-color: #3b82f6;
        }

        QLineEdit:focus {
            border-color: #3b82f6;
            background-color: #f8fafc;
        }

        QLineEdit:hover {
            border-color: #94a3b8;
        }

        /* Modern Buttons */
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #3b82f6, stop:1 #1d4ed8);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 600;
            font-size: 10pt;
            min-height: 20px;
        }

        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #2563eb, stop:1 #1e40af);
        }

        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #1d4ed8, stop:1 #1e3a8a);
        }

        /* Special Button Styles */
        QPushButton[objectName="yeniUrunBtn"] {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #10b981, stop:1 #059669);
        }

        QPushButton[objectName="yeniUrunBtn"]:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #059669, stop:1 #047857);
        }

        QPushButton[objectName="menuButton"] {
            background: #f8fafc;
            color: #64748b;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            font-weight: 600;
            font-size: 14pt;
        }

        QPushButton[objectName="menuButton"]:hover {
            background: #f1f5f9;
            color: #475569;
            border-color: #cbd5e1;
        }

        QPushButton[objectName="degistirBtn"] {
            background: transparent;
            color: #3b82f6;
            text-decoration: underline;
            padding: 4px 8px;
        }

        QPushButton[objectName="degistirBtn"]:hover {
            color: #1d4ed8;
        }

        /* Search Container */
        QFrame#searchContainer {
            background-color: white;
            border: 2px solid #e2e8f0;
            border-radius: 12px;
            margin: 5px;
        }

        QLineEdit#searchInput {
            border: none;
            background: transparent;
            font-size: 11pt;
            padding: 8px 0px;
        }

        QPushButton#clearSearchBtn {
            background: #f1f5f9;
            color: #64748b;
            border: none;
            border-radius: 6px;
            font-weight: 600;
        }

        QPushButton#clearSearchBtn:hover {
            background: #e2e8f0;
            color: #475569;
        }

        QPushButton#filterBtn {
            background: #f8fafc;
            color: #475569;
            border: 2px solid #e2e8f0;
        }

        QPushButton#filterBtn:hover {
            background: #f1f5f9;
            border-color: #cbd5e1;
        }

        /* Modern Table */
        QTableWidget {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            gridline-color: #f1f5f9;
            font-size: 10pt;
            selection-background-color: #dbeafe;
            alternate-background-color: #f8fafc;
        }

        QTableWidget::item {
            padding: 16px 12px;
            border: none;
        }

        QTableWidget::item:hover {
            background-color: #f1f5f9;
        }

        QTableWidget::item:selected {
            background-color: #dbeafe;
            color: #1e40af;
        }

        /* Table Header */
        QHeaderView::section {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #f8fafc, stop:1 #f1f5f9);
            color: #475569;
            padding: 16px 12px;
            border: none;
            border-bottom: 2px solid #e2e8f0;
            font-weight: 600;
            font-size: 9pt;
        }

        QHeaderView::section:horizontal {
            border-right: 1px solid #e2e8f0;
        }

        /* Form Styling */
        QFrame#eklemeFormu {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 24px;
            margin: 10px;
        }

        QFormLayout {
            spacing: 20px;
        }

        QLabel {
            color: #374151;
            font-weight: 600;
            font-size: 10pt;
        }

        /* Metric Cards */
        QFrame#metricCard {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 8px;
            margin: 4px;
        }

        /* Modern Menu */
        QMenuBar {
            background: white;
            border-bottom: 1px solid #e2e8f0;
            padding: 8px;
        }

        QMenuBar::item {
            padding: 8px 16px;
            border-radius: 6px;
            margin: 2px;
        }

        QMenuBar::item:selected {
            background-color: #f1f5f9;
            color: #1e40af;
        }

Main
        QMenu {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 8px;
        }

        QMenu::item {
            padding: 10px 20px;
            border-radius: 4px;
            margin: 2px;
        }

        QMenu::item:selected {
            background-color: #f1f5f9;
            color: #1e40af;
        }

        QMenu::separator {
            height: 1px;
            background-color: #e2e8f0;
            margin: 8px 0;
        }

        QMenu::item[objectName="dangerAction"]:selected {
            background-color: #fef2f2;
            color: #dc2626;
        }

        /* Status Bar */
        QStatusBar {
            background-color: white;
            color: #64748b;
            border-top: 1px solid #e2e8f0;
            padding: 8px 16px;
            font-size: 9pt;
        }

        /* Dialog Styling */
        QDialog {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #f8fafc, stop:1 #f1f5f9);
        }

        /* Scrollbar Styling */
        QScrollBar:vertical {
            background: #f1f5f9;
            width: 12px;
            border-radius: 6px;
        }

        QScrollBar::handle:vertical {
            background: #cbd5e1;
            border-radius: 6px;
            min-height: 20px;
        }

        QScrollBar::handle:vertical:hover {
            background: #94a3b8;
        }

        QScrollBar:horizontal {
            background: #f1f5f9;
            height: 12px;
            border-radius: 6px;
        }

        QScrollBar::handle:horizontal {
            background: #cbd5e1;
            border-radius: 6px;
            min-width: 20px;
        }

        QScrollBar::handle:horizontal:hover {
            background: #94a3b8;
        }
    """)

    kontrolcu = AnaKontrolcu()
    kontrolcu.baslat()
    sys.exit(app.exec_())