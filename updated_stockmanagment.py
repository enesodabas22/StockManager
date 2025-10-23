import sys
import sqlite3
import hashlib
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QMessageBox,
                             QTableWidget, QTableWidgetItem, QMenu, QAction,
                             QInputDialog, QHeaderView, QFrame, QMainWindow,
                             QDialog, QFormLayout, QDialogButtonBox,
                             QAbstractItemView)  # Seçim davranışını ayarlamak için eklendi
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QFont, QColor  # QColor eklendi


# --- Veritabanı Sınıfı (ID getir eklendi) ---
class VeritabaniYoneticisi:
    def __init__(self, db_adi="stok_veritabani.db"):
        self.baglanti = sqlite3.connect(db_adi)
        self.cursor = self.baglanti.cursor()
        self.tablolari_olustur()

    def tablolari_olustur(self):
        # ID sütunu zaten PRIMARY KEY olarak vardı.
        self.cursor.execute(
            "CREATE TABLE IF NOT EXISTS urunler (id INTEGER PRIMARY KEY, ad TEXT NOT NULL UNIQUE, kategori TEXT, fiyat REAL NOT NULL DEFAULT 0.0, miktar INTEGER NOT NULL)")
        self.cursor.execute(
            "CREATE TABLE IF NOT EXISTS kullanicilar (id INTEGER PRIMARY KEY, kullanici_adi TEXT NOT NULL UNIQUE, sifre_hash TEXT NOT NULL)")
        self.baglanti.commit()

    def urunleri_getir(self):
        # YENİLİK: ID'yi de seçiyoruz
        self.cursor.execute("SELECT id, ad, kategori, fiyat, miktar FROM urunler ORDER BY ad ASC")
        return self.cursor.fetchall()

    def urun_detay_getir(self, ad):
        self.cursor.execute("SELECT id, ad, kategori, fiyat, miktar FROM urunler WHERE ad = ?", (ad,))
        return self.cursor.fetchone()  # ID'yi de döndürür ama düzenleme dialoğu ID'yi kullanmıyor şimdilik

    def urun_ekle(self, ad, kategori, fiyat, miktar):
        try:
            self.cursor.execute("INSERT INTO urunler (ad, kategori, fiyat, miktar) VALUES (?, ?, ?, ?)",
                                (ad, kategori, fiyat, miktar))
        except sqlite3.IntegrityError:
            self.cursor.execute("UPDATE urunler SET miktar = miktar + ? WHERE ad = ?", (miktar, ad))
        self.baglanti.commit()

    def urun_detay_guncelle(self, eski_ad, yeni_ad, kategori, fiyat):
        self.cursor.execute("UPDATE urunler SET ad = ?, kategori = ?, fiyat = ? WHERE ad = ?",
                            (yeni_ad, kategori, fiyat, eski_ad)); self.baglanti.commit()

    def kullanici_sayisi_getir(self):
        self.cursor.execute("SELECT COUNT(*) FROM kullanicilar"); return self.cursor.fetchone()[0]

    def sifre_hashle(self, s):
        return hashlib.sha256(s.encode()).hexdigest()

    def kullanici_ekle(self, k_adi, s):
        try:
            self.cursor.execute("INSERT INTO kullanicilar (kullanici_adi, sifre_hash) VALUES (?, ?)", (
            k_adi, self.sifre_hashle(s))); self.baglanti.commit(); return True, "Kullanıcı oluşturuldu."
        except sqlite3.IntegrityError:
            return False, "Bu kullanıcı adı zaten alınmış."

    def kullanici_dogrula(self, k_adi, s):
        self.cursor.execute("SELECT * FROM kullanicilar WHERE kullanici_adi = ? AND sifre_hash = ?",
                            (k_adi, self.sifre_hashle(s))); return self.cursor.fetchone() is not None

    def kullanici_bilgilerini_guncelle(self, e_kadi, y_kadi, y_sifre):
        try:
            self.cursor.execute("UPDATE kullanicilar SET kullanici_adi = ?, sifre_hash = ? WHERE kullanici_adi = ?", (
            y_kadi, self.sifre_hashle(y_sifre), e_kadi)); self.baglanti.commit(); return True, "Bilgiler güncellendi."
        except sqlite3.IntegrityError:
            return False, "Yeni kullanıcı adı başkası tarafından kullanılıyor."

    def urun_guncelle(self, ad, miktar_farki):
        self.cursor.execute("UPDATE urunler SET miktar = miktar + ? WHERE ad = ?",
                            (miktar_farki, ad)); self.baglanti.commit()

    def urun_sil(self, ad):
        self.cursor.execute("DELETE FROM urunler WHERE ad = ?", (ad,)); self.baglanti.commit()

    def mevcut_miktar_getir(self, ad):
        self.cursor.execute("SELECT miktar FROM urunler WHERE ad = ?", (ad,)); sonuc = self.cursor.fetchone(); return \
        sonuc[0] if sonuc else 0


# --- Arayüz Sınıfları ---
class UrunDuzenlemeDialog(QDialog):
    # Bu pencere şimdilik ID kullanmıyor, isim üzerinden devam ediyor.
    def __init__(self, urun_detaylari, parent=None):
        super().__init__(parent);
        self.setWindowTitle("Ürün Bilgilerini Düzenle")
        self.form_layout = QFormLayout(self)
        self.ad_input = QLineEdit(urun_detaylari[1]);  # Index 1 artık Ad
        self.kategori_input = QLineEdit(urun_detaylari[2]);  # Index 2 Kategori
        self.fiyat_input = QLineEdit(str(urun_detaylari[3]))  # Index 3 Fiyat
        self.form_layout.addRow("Ürün Adı:", self.ad_input);
        self.form_layout.addRow("Kategori:", self.kategori_input);
        self.form_layout.addRow("Fiyat (₺):", self.fiyat_input)
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel);
        self.buttonBox.accepted.connect(self.accept);
        self.buttonBox.rejected.connect(self.reject)
        self.form_layout.addWidget(self.buttonBox)

    def get_data(self): return (
    self.ad_input.text(), self.kategori_input.text(), float(self.fiyat_input.text().replace(',', '.')))


class StokUygulamasiWidget(QWidget):
    def __init__(self, veritabani_yoneticisi, status_bar):
        super().__init__();
        self.veritabani = veritabani_yoneticisi;
        self.status_bar = status_bar
        self.arayuz_olustur();
        self.stogu_guncelle_arayuz()

    def arayuz_olustur(self):
        ana_duzen = QVBoxLayout(self);
        ana_duzen.setContentsMargins(15, 15, 15, 15);
        ana_duzen.setSpacing(10)
        ust_duzen = QHBoxLayout()
        self.arama_input = QLineEdit();
        self.arama_input.setPlaceholderText("Ürün adında ara...");
        self.arama_input.textChanged.connect(self.tabloyu_filtrele);
        ust_duzen.addWidget(self.arama_input, 1)
        self.yeni_urun_goster_btn = QPushButton(QIcon.fromTheme("list-add"), " Yeni Ürün Ekle");
        self.yeni_urun_goster_btn.setObjectName("yeniUrunBtn");
        self.yeni_urun_goster_btn.clicked.connect(self.yeni_urun_formu_goster_gizle);
        ust_duzen.addWidget(self.yeni_urun_goster_btn)
        ana_duzen.addLayout(ust_duzen)

        self.ekleme_formu_frame = QFrame();
        self.ekleme_formu_frame.setObjectName("eklemeFormu");
        self.ekleme_formu_frame.setFrameShape(QFrame.StyledPanel)
        ekleme_duzen = QFormLayout(self.ekleme_formu_frame);
        ekleme_duzen.setContentsMargins(10, 10, 10, 10);
        ekleme_duzen.setSpacing(8)
        self.yeni_urun_input = QLineEdit();
        self.yeni_kategori_input = QLineEdit();
        self.yeni_fiyat_input = QLineEdit();
        self.yeni_miktar_input = QLineEdit()
        self.onayla_ekle_btn = QPushButton(QIcon.fromTheme("dialog-ok"), " Onayla");
        self.onayla_ekle_btn.setDefault(True)
        ekleme_duzen.addRow("Ürün Adı:", self.yeni_urun_input);
        ekleme_duzen.addRow("Kategori:", self.yeni_kategori_input);
        ekleme_duzen.addRow("Fiyat (₺):", self.yeni_fiyat_input);
        ekleme_duzen.addRow("Miktar:", self.yeni_miktar_input);
        ekleme_duzen.addRow(self.onayla_ekle_btn)
        self.ekleme_formu_frame.hide();
        ana_duzen.addWidget(self.ekleme_formu_frame)
        self.onayla_ekle_btn.clicked.connect(self.yeni_urun_ekle)

        self.stok_tablosu = QTableWidget()
        # YENİLİK: Sütun sayısı 6 oldu (ID eklendi)
        self.stok_tablosu.setColumnCount(6)
        self.stok_tablosu.setHorizontalHeaderLabels(["ID", "Ürün Adı", "Kategori", "Fiyat", "Miktar", "İşlemler"])
        # Sütun genişlik ayarları
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # ID dar olsun
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # Ürün Adı genişlesin
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)  # İşlemler dar olsun
        self.stok_tablosu.setAlternatingRowColors(True)
        self.stok_tablosu.setShowGrid(False)
        self.stok_tablosu.verticalHeader().setVisible(False)
        self.stok_tablosu.setSelectionBehavior(QAbstractItemView.SelectRows)  # Satır bazında seçim
        self.stok_tablosu.setSelectionMode(QAbstractItemView.SingleSelection)  # Tek satır seçimi
        self.stok_tablosu.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Tabloyu düzenlenemez yap
        ana_duzen.addWidget(self.stok_tablosu)

    def stogu_guncelle_arayuz(self):
        self.stok_tablosu.setRowCount(0);
        urun_listesi = self.veritabani.urunleri_getir();
        self.stok_tablosu.setRowCount(len(urun_listesi))

        # YENİLİK: Font tanımlamaları
        urun_fontu = QFont("Segoe UI Semibold", 11);  # Kalın ve biraz büyük
        normal_font = QFont("Segoe UI", 10);  # Normal boyut
        id_fontu = QFont("Segoe UI", 9);  # ID için daha küçük

        id_renk = QColor("#A0A0A0")  # ID için soluk renk (gri tonu)

        for satir, (id_val, ad, kategori, fiyat, miktar) in enumerate(urun_listesi):
            # ID hücresi
            id_item = QTableWidgetItem(str(id_val))
            id_item.setFont(id_fontu)
            id_item.setForeground(id_renk)
            id_item.setTextAlignment(Qt.AlignCenter)

            # Ürün Adı hücresi
            ad_item = QTableWidgetItem(ad);
            ad_item.setFont(urun_fontu)

            # Diğer hücreler
            kategori_item = QTableWidgetItem(kategori);
            kategori_item.setFont(normal_font)
            fiyat_item = QTableWidgetItem(f"{fiyat:.2f} ₺");
            fiyat_item.setFont(normal_font);
            fiyat_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            miktar_item = QTableWidgetItem(str(miktar));
            miktar_item.setFont(normal_font);
            miktar_item.setTextAlignment(Qt.AlignCenter)

            self.stok_tablosu.setItem(satir, 0, id_item)  # ID eklendi (index 0)
            self.stok_tablosu.setItem(satir, 1, ad_item)  # Diğerleri bir kaydırıldı
            self.stok_tablosu.setItem(satir, 2, kategori_item)
            self.stok_tablosu.setItem(satir, 3, fiyat_item)
            self.stok_tablosu.setItem(satir, 4, miktar_item)

            menu_btn = QPushButton("...");
            menu_btn.setFixedWidth(40);
            menu_btn.setObjectName("menuButton");
            menu_btn.clicked.connect(lambda ch, r=satir: self.guncelle_menusu_goster(r, menu_btn));
            self.stok_tablosu.setCellWidget(satir, 5, menu_btn)  # Index 5
        self.stok_tablosu.resizeRowsToContents()

    def guncelle_menusu_goster(self, satir_indeksi, buton):
        # Ürün adını ID yerine kullanmaya devam ediyoruz şimdilik (isimler unique)
        urun_adi = self.stok_tablosu.item(satir_indeksi, 1).text()  # Index 1 artık Ad
        menu = QMenu(self)
        duzenle = QAction(QIcon.fromTheme("document-edit"), "Bilgileri Düzenle", self);
        duzenle.triggered.connect(lambda: self.urun_duzenle(urun_adi));
        menu.addAction(duzenle)
        menu.addSeparator()
        artir = QAction(QIcon.fromTheme("go-up"), "Stok Artır/Ekle", self);
        artir.triggered.connect(lambda: self.miktar_girdi_goster(urun_adi, 'artır'));
        menu.addAction(artir)
        eksilt = QAction(QIcon.fromTheme("go-down"), "Stok Eksilt/Çıkar", self);
        eksilt.triggered.connect(lambda: self.miktar_girdi_goster(urun_adi, 'eksilt'));
        menu.addAction(eksilt)
        menu.addSeparator()
        sil = QAction(QIcon.fromTheme("edit-delete"), "Ürünü Sil", self);
        sil.triggered.connect(lambda: self.urun_sil(urun_adi));
        menu.addAction(sil)
        menu.exec_(buton.mapToGlobal(buton.rect().bottomLeft()))

    def urun_duzenle(self, urun_adi):
        urun_detaylari = self.veritabani.urun_detay_getir(urun_adi)  # Bu artık ID'yi de içeriyor ama dialog kullanmıyor
        dialog = UrunDuzenlemeDialog(urun_detaylari, self)
        if dialog.exec_() == QDialog.Accepted:
            yeni_ad, yeni_kategori, yeni_fiyat = dialog.get_data()
            self.veritabani.urun_detay_guncelle(urun_adi, yeni_ad, yeni_kategori, yeni_fiyat);
            self.stogu_guncelle_arayuz();
            self.status_bar.showMessage(f"'{yeni_ad}' güncellendi.", 3000)

    # ... (Diğer metodlar aynı kaldı) ...
    def yeni_urun_ekle(self):
        ad = self.yeni_urun_input.text().strip().capitalize();
        kat = self.yeni_kategori_input.text().strip();
        f_str = self.yeni_fiyat_input.text().strip();
        m_str = self.yeni_miktar_input.text().strip()
        if not all([ad, f_str, m_str]): QMessageBox.warning(self, "Uyarı",
                                                            "Ürün Adı, Fiyat ve Miktar zorunludur."); return
        try:
            f = float(f_str.replace(',', '.')); m = int(m_str); assert m > 0 and f >= 0
        except(ValueError, AssertionError):
            QMessageBox.warning(self, "Hata", "Fiyat ve Miktar pozitif sayı olmalıdır."); return
        self.veritabani.urun_ekle(ad, kat, f, m);
        self.stogu_guncelle_arayuz();
        self.yeni_urun_formu_goster_gizle()
        self.status_bar.showMessage(f"'{ad}' eklendi/güncellendi.", 3000)
        for item in [self.yeni_urun_input, self.yeni_kategori_input, self.yeni_fiyat_input,
                     self.yeni_miktar_input]: item.clear()

    def tabloyu_filtrele(self, metin):
        for i in range(self.stok_tablosu.rowCount()): self.stok_tablosu.setRowHidden(i,
                                                                                     metin.lower() not in self.stok_tablosu.item(
                                                                                         i,
                                                                                         1).text().lower())  # Index 1 Ad

    def yeni_urun_formu_goster_gizle(self):
        if self.ekleme_formu_frame.isVisible():
            self.ekleme_formu_frame.hide()
        else:
            self.ekleme_formu_frame.show(); self.yeni_urun_input.setFocus()

    def miktar_girdi_goster(self, u_adi, mod):
        fiil = "Artır" if mod == 'artır' else "Eksilt";
        m, ok = QInputDialog.getInt(self, f"Stoğu {fiil}", f"Lütfen miktarı girin:", 1, 1, 999999)
        if ok: self.stok_miktari_degistir(u_adi, m if mod == 'artır' else -m)

    def stok_miktari_degistir(self, u_adi, m_farki):
        mevcut = self.veritabani.mevcut_miktar_getir(u_adi)
        if mevcut + m_farki < 0: QMessageBox.warning(self, "Uyarı", "Stok eksiye düşemez."); return
        self.veritabani.urun_guncelle(u_adi, m_farki);
        self.stogu_guncelle_arayuz();
        self.status_bar.showMessage(f"'{u_adi}' stoğu güncellendi.", 3000)

    def urun_sil(self, u_adi):
        onay = QMessageBox.question(self, "Silmeyi Onayla", f"Emin misiniz?", QMessageBox.Yes | QMessageBox.No,
                                    QMessageBox.No)
        if onay == QMessageBox.Yes: self.veritabani.urun_sil(
            u_adi); self.stogu_guncelle_arayuz(); self.status_bar.showMessage(f"'{u_adi}' silindi.", 3000)


class AnaPencere(QMainWindow):
    def __init__(self, kullanici_adi, veritabani_yoneticisi):
        super().__init__()
        self.setWindowTitle(f"Stok Yönetimi - [{kullanici_adi}]");
        self.setGeometry(100, 100, 1000, 800)  # Biraz daha büyütüldü
        self.status_bar = self.statusBar();
        self.status_bar.showMessage("Hoş geldiniz!", 3000)
        self.merkezi_widget = StokUygulamasiWidget(veritabani_yoneticisi, self.status_bar);
        self.setCentralWidget(self.merkezi_widget)


# --- Giriş ve Kontrolcü Sınıfları (Aynı kaldı) ---
class IlkKurulumPenceresi(QWidget):
    kurulum_tamamlandi = pyqtSignal(str)

    def __init__(self, veritabani_yoneticisi):
        super().__init__();
        self.veritabani = veritabani_yoneticisi;
        self.setWindowTitle("İlk Kurulum");
        self.setGeometry(400, 400, 400, 250)
        duzen = QVBoxLayout(self);
        duzen.addWidget(QLabel("<h2>Yönetici Hesabı Oluştur</h2>"))
        self.k_adi = QLineEdit();
        self.k_adi.setPlaceholderText("Kullanıcı Adı")
        self.sifre = QLineEdit();
        self.sifre.setPlaceholderText("Şifre");
        self.sifre.setEchoMode(QLineEdit.Password)
        self.sifre_t = QLineEdit();
        self.sifre_t.setPlaceholderText("Şifre Tekrar");
        self.sifre_t.setEchoMode(QLineEdit.Password)
        btn = QPushButton("Hesabı Oluştur");
        btn.setDefault(True);
        btn.clicked.connect(self.hesap_olustur)
        duzen.addWidget(self.k_adi);
        duzen.addWidget(self.sifre);
        duzen.addWidget(self.sifre_t);
        duzen.addWidget(btn)

    def hesap_olustur(self):
        k, s, st = self.k_adi.text().strip(), self.sifre.text(), self.sifre_t.text()
        if not k or not s: QMessageBox.warning(self, "Hata", "Alanlar boş bırakılamaz."); return
        if s != st: QMessageBox.warning(self, "Hata", "Şifreler eşleşmiyor."); return
        b, m = self.veritabani.kullanici_ekle(k, s)
        if b:
            self.kurulum_tamamlandi.emit(k); self.close()
        else:
            QMessageBox.critical(self, "Hata", m)


class KullaniciDegistirPenceresi(QWidget):
    degisiklik_yapildi = pyqtSignal()

    def __init__(self, veritabani_yoneticisi):
        super().__init__();
        self.veritabani = veritabani_yoneticisi;
        self.setWindowTitle("Bilgileri Değiştir");
        self.setGeometry(400, 300, 450, 400)
        duzen = QVBoxLayout(self);
        duzen.addWidget(QLabel("<h3>Mevcut Bilgileri Doğrula</h3>"))
        self.e_kadi = QLineEdit();
        self.e_kadi.setPlaceholderText("Mevcut Kullanıcı Adı")
        self.e_sifre = QLineEdit();
        self.e_sifre.setPlaceholderText("Mevcut Şifre");
        self.e_sifre.setEchoMode(QLineEdit.Password)
        duzen.addWidget(self.e_kadi);
        duzen.addWidget(self.e_sifre);
        duzen.addWidget(QLabel("<h3>Yeni Bilgileri Girin</h3>"))
        self.y_kadi = QLineEdit();
        self.y_kadi.setPlaceholderText("Yeni Kullanıcı Adı")
        self.y_sifre = QLineEdit();
        self.y_sifre.setPlaceholderText("Yeni Şifre");
        self.y_sifre.setEchoMode(QLineEdit.Password)
        self.y_sifre_t = QLineEdit();
        self.y_sifre_t.setPlaceholderText("Yeni Şifre Tekrar");
        self.y_sifre_t.setEchoMode(QLineEdit.Password)
        duzen.addWidget(self.y_kadi);
        duzen.addWidget(self.y_sifre);
        duzen.addWidget(self.y_sifre_t)
        btn = QPushButton("Değişiklikleri Onayla");
        btn.setDefault(True);
        btn.clicked.connect(self.bilgileri_degistir)
        duzen.addStretch();
        duzen.addWidget(btn)

    def bilgileri_degistir(self):
        ek, es = self.e_kadi.text().strip(), self.e_sifre.text();
        yk, ys, yst = self.y_kadi.text().strip(), self.y_sifre.text(), self.y_sifre_t.text()
        if not all([ek, es, yk, ys]): QMessageBox.warning(self, "Hata", "Tüm alanlar doldurulmalıdır."); return
        if ys != yst: QMessageBox.warning(self, "Hata", "Yeni şifreler eşleşmiyor."); return
        if not self.veritabani.kullanici_dogrula(ek, es): QMessageBox.warning(self, "Hata",
                                                                              "Mevcut bilgiler yanlış."); return
        b, m = self.veritabani.kullanici_bilgilerini_guncelle(ek, yk, ys)
        if b:
            QMessageBox.information(self, "Başarılı", m); self.degisiklik_yapildi.emit(); self.close()
        else:
            QMessageBox.warning(self, "Hata", m)


class GirisPenceresi(QWidget):
    login_basarili = pyqtSignal(str);
    degistirme_penceresi_iste = pyqtSignal()

    def __init__(self, veritabani_yoneticisi):
        super().__init__();
        self.veritabani = veritabani_yoneticisi;
        self.setWindowTitle("Giriş");
        self.setGeometry(400, 400, 400, 250)
        duzen = QVBoxLayout(self);
        duzen.addWidget(QLabel("<h2>Giriş Yap</h2>"))
        self.k_adi = QLineEdit();
        self.k_adi.setPlaceholderText("Kullanıcı Adı")
        self.sifre = QLineEdit();
        self.sifre.setPlaceholderText("Şifre");
        self.sifre.setEchoMode(QLineEdit.Password)
        login_btn = QPushButton("Giriş Yap");
        login_btn.setDefault(True);
        login_btn.clicked.connect(self.login_kontrol)
        degistir_btn = QPushButton("Bilgileri Değiştir");
        degistir_btn.setObjectName("degistirBtn");
        degistir_btn.setFlat(True);
        degistir_btn.clicked.connect(self.degistirme_penceresi_iste.emit)
        duzen.addWidget(self.k_adi);
        duzen.addWidget(self.sifre);
        duzen.addWidget(login_btn);
        duzen.addWidget(degistir_btn, 0, Qt.AlignRight)
        self.sifre.returnPressed.connect(self.login_kontrol)

    def login_kontrol(self):
        k, s = self.k_adi.text(), self.sifre.text()
        if self.veritabani.kullanici_dogrula(k, s):
            self.login_basarili.emit(k); self.close()
        else:
            QMessageBox.warning(self, "Hata", "Kullanıcı adı veya şifre hatalı.")


class AnaKontrolcu:
    def __init__(self):
        self.veritabani = VeritabaniYoneticisi();
        self.mevcut_pencere = None;
        self.ana_pencere = None;
        self.degistirme_penceresi = None

    def baslat(self):
        if self.veritabani.kullanici_sayisi_getir() == 0:
            self.mevcut_pencere = IlkKurulumPenceresi(self.veritabani);
            self.mevcut_pencere.kurulum_tamamlandi.connect(self.ana_pencereyi_goster)
        else:
            self.mevcut_pencere = GirisPenceresi(self.veritabani);
            self.mevcut_pencere.login_basarili.connect(self.ana_pencereyi_goster)
            self.mevcut_pencere.degistirme_penceresi_iste.connect(self.degistirme_penceresini_goster)
        self.mevcut_pencere.show()

    def ana_pencereyi_goster(self, kullanici_adi):
        self.ana_pencere = AnaPencere(kullanici_adi, self.veritabani);
        self.ana_pencere.show()
        if self.mevcut_pencere: self.mevcut_pencere.close(); self.mevcut_pencere = None

    def degistirme_penceresini_goster(self):
        self.degistirme_penceresi = KullaniciDegistirPenceresi(self.veritabani);
        self.degistirme_penceresi.degisiklik_yapildi.connect(self.login_penceresine_don)
        self.degistirme_penceresi.show();
        if self.mevcut_pencere: self.mevcut_pencere.hide()

    def login_penceresine_don(self):
        if self.degistirme_penceresi: self.degistirme_penceresi.close()
        if self.mevcut_pencere: self.mevcut_pencere.show()


# --- Uygulama Başlatma ve Daha Gelişmiş Stil ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Geliştirilmiş Koyu Tema (Nord Teması Benzeri)
    app.setStyleSheet("""
        QWidget { 
            background-color: #2E3440; /* Ana arka plan */
            color: #ECEFF4; /* Ana metin rengi */
            font-family: 'Segoe UI Semibold', 'Segoe UI', sans-serif; /* Öncelikli Semibold, yoksa Segoe UI */
            font-size: 11pt; /* Genel font boyutu biraz büyütüldü */
        }
        QMainWindow, QDialog { 
            border: 1px solid #4C566A; /* Pencerelere hafif kenarlık */
        } 
        QLineEdit, QDialog { 
            background-color: #3B4252; /* Giriş alanları arka plan */
            border: 1px solid #4C566A; 
            border-radius: 5px; /* Daha yuvarlak kenarlar */
            padding: 8px; /* Daha fazla iç boşluk */
            color: #ECEFF4; 
        }
        QLineEdit:focus {
            border: 1px solid #88C0D0; /* Odaklanınca kenarlık rengi */
        }
        QTableWidget { 
            background-color: #3B4252; 
            border: 1px solid #4C566A; 
            border-radius: 4px; 
            gridline-color: #434C5E; /* Izgara çizgisi rengi */
            font-size: 10pt; /* Tablo içeriği biraz daha küçük */
        }
        QTableWidget::item { 
            padding: 10px 8px; /* Hücre içi boşluk (Yükseklik, Genişlik) */
            border-bottom: 1px solid #434C5E; 
        }
        QTableWidget::item:alternate { 
            background-color: #434C5E; /* Zebra deseni */
        }
        QTableWidget::item:hover { 
            background-color: #4C566A; /* Satır üzerine gelince */
        }
        QTableWidget::item:selected { 
             background-color: #5E81AC; /* Seçili satır rengi */
             color: #ECEFF4;
        }
        QHeaderView::section { 
            background-color: #434C5E; 
            padding: 8px 5px; /* Başlık iç boşluğu */
            border: none; 
            border-bottom: 1px solid #4C566A; /* Başlık alt çizgisi */
            font-weight: bold; 
            font-size: 10pt;
        }
        QHeaderView::section:vertical { /* Sol dikey başlık (numaralar) */
             background-color: #434C5E;
             border: none;
        } 
        QPushButton { 
            background-color: #5E81AC; /* Buton rengi (mavi tonu) */
            color: #ECEFF4;
            border: none; 
            padding: 9px 18px; /* Buton iç boşluğu */
            border-radius: 5px; 
            font-weight: bold; 
        }
        QPushButton:hover { 
            background-color: #81A1C1; /* Üzerine gelince daha açık mavi */
        }
        QPushButton:pressed { 
            background-color: #88C0D0; /* Basılınca en açık mavi */
        }
        QPushButton:disabled { 
            background-color: #4C566A; 
            color: #6a748b;
        }
        QPushButton[objectName="menuButton"] { /* Menü butonu (...) */
            padding: 4px 10px; 
            font-weight: bold;
            background-color: #4C566A;
        }
        QPushButton[objectName="yeniUrunBtn"] { /* Yeni Ürün Ekle butonu */
             background-color: #A3BE8C; /* Yeşil tonu */
        }
        QPushButton[objectName="yeniUrunBtn"]:hover { background-color: #b4d0a0; }
        QPushButton[objectName="degistirBtn"] { /* Link Görünümlü Buton */
            background-color: transparent; 
            color: #88C0D0; 
            font-weight: normal; 
            text-align: right; 
            padding: 0px; 
            border: none;
        }
        QPushButton[objectName="degistirBtn"]:hover { color: #5E81AC; }
        QStatusBar { color: #D8DEE9; padding: 3px; }
        QMenu { background-color: #3B4252; border: 1px solid #4C566A; padding: 5px; }
        QMenu::item { padding: 6px 22px; }
        QMenu::item:selected { background-color: #5E81AC; }
        QMenu::separator { height: 1px; background-color: #4C566A; margin: 4px 0px; }
        QFrame#eklemeFormu { background-color: #434C5E; border-radius: 5px; border: 1px solid #4C566A;} 
    """)
    kontrolcu = AnaKontrolcu()
    kontrolcu.baslat()
    sys.exit(app.exec_())
