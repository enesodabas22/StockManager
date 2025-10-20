import sys
import sqlite3
import hashlib
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QMessageBox,
                             QTableWidget, QTableWidgetItem, QMenu, QAction,
                             QInputDialog, QHeaderView, QFrame, QMainWindow,
                             QDialog, QFormLayout, QDialogButtonBox)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon

# --- Veritabanı yöneticisi ---
class VeritabaniYoneticisi:
    """SQLite veritabanı işlemlerini yönetir."""
    def __init__(self, db_adi="stok_veritabani.db"):
        self.baglanti = sqlite3.connect(db_adi)
        self.cursor = self.baglanti.cursor()
        self.tablolari_olustur()
        # Kullanılmayan ikonları yüklememek için statik tema ayarı
        QApplication.setQuitOnLastWindowClosed(True)

    def tablolari_olustur(self):
        # Ürünler tablosu
        self.cursor.execute(
            "CREATE TABLE IF NOT EXISTS urunler (id INTEGER PRIMARY KEY, ad TEXT NOT NULL UNIQUE, kategori TEXT, fiyat REAL NOT NULL DEFAULT 0.0, miktar INTEGER NOT NULL)"
        )
        # Kullanıcılar tablosu (sha256 hash ile şifre saklanır)
        self.cursor.execute(
            "CREATE TABLE IF NOT EXISTS kullanicilar (id INTEGER PRIMARY KEY, kullanici_adi TEXT NOT NULL UNIQUE, sifre_hash TEXT NOT NULL)"
        )
        self.baglanti.commit()

    def urunleri_getir(self):
        self.cursor.execute("SELECT ad, kategori, fiyat, miktar FROM urunler ORDER BY ad ASC")
        return self.cursor.fetchall()

    def urun_detay_getir(self, ad):
        self.cursor.execute("SELECT ad, kategori, fiyat, miktar FROM urunler WHERE ad = ?", (ad,))
        return self.cursor.fetchone()

    def urun_ekle(self, ad, kategori, fiyat, miktar):
        try:
            # Yeni ürün ekle
            self.cursor.execute(
                "INSERT INTO urunler (ad, kategori, fiyat, miktar) VALUES (?, ?, ?, ?)",
                (ad, kategori, fiyat, miktar),
            )
        except sqlite3.IntegrityError:
            # Eğer aynı isim varsa miktarı arttır
            self.cursor.execute("UPDATE urunler SET miktar = miktar + ? WHERE ad = ?", (miktar, ad))
        self.baglanti.commit()

    def urun_detay_guncelle(self, eski_ad, yeni_ad, kategori, fiyat):
        self.cursor.execute(
            "UPDATE urunler SET ad = ?, kategori = ?, fiyat = ? WHERE ad = ?",
            (yeni_ad, kategori, fiyat, eski_ad),
        )
        self.baglanti.commit()

    def kullanici_sayisi_getir(self):
        self.cursor.execute("SELECT COUNT(*) FROM kullanicilar")
        return self.cursor.fetchone()[0]

    def sifre_hashle(self, s):
        """Şifreyi SHA256 ile hashler."""
        return hashlib.sha256(s.encode()).hexdigest()

    def kullanici_ekle(self, k_adi, s):
        try:
            self.cursor.execute(
                "INSERT INTO kullanicilar (kullanici_adi, sifre_hash) VALUES (?, ?)",
                (k_adi, self.sifre_hashle(s)),
            )
            self.baglanti.commit()
            return True, "Kullanıcı oluşturuldu."
        except sqlite3.IntegrityError:
            return False, "Bu kullanıcı adı zaten alınmış."

    def kullanici_dogrula(self, k_adi, s):
        self.cursor.execute(
            "SELECT * FROM kullanicilar WHERE kullanici_adi = ? AND sifre_hash = ?",
            (k_adi, self.sifre_hashle(s)),
        )
        return self.cursor.fetchone() is not None

    def kullanici_bilgilerini_guncelle(self, e_kadi, y_kadi, y_sifre):
        try:
            self.cursor.execute(
                "UPDATE kullanicilar SET kullanici_adi = ?, sifre_hash = ? WHERE kullanici_adi = ?",
                (y_kadi, self.sifre_hashle(y_sifre), e_kadi),
            )
            self.baglanti.commit()
            return True, "Bilgiler güncellendi."
        except sqlite3.IntegrityError:
            return False, "Yeni kullanıcı adı başkası tarafından kullanılıyor."

    def urun_guncelle(self, ad, miktar_farki):
        self.cursor.execute("UPDATE urunler SET miktar = miktar + ? WHERE ad = ?", (miktar_farki, ad))
        self.baglanti.commit()

    def urun_sil(self, ad):
        self.cursor.execute("DELETE FROM urunler WHERE ad = ?", (ad,))
        self.baglanti.commit()

    def mevcut_miktar_getir(self, ad):
        self.cursor.execute("SELECT miktar FROM urunler WHERE ad = ?", (ad,))
        sonuc = self.cursor.fetchone()
        return sonuc[0] if sonuc else 0


# --- Ürün Düzenleme Dialogu ---
class UrunDuzenlemeDialog(QDialog):
    """Ürün detaylarını düzenlemek için kullanılan diyalog."""
    def __init__(self, urun_detaylari, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ürün Bilgilerini Düzenle")
        self.form_layout = QFormLayout(self)
        self.ad_input = QLineEdit(urun_detaylari[0])
        self.kategori_input = QLineEdit(urun_detaylari[1])
        # Fiyatı string olarak alıyoruz
        self.fiyat_input = QLineEdit(f"{urun_detaylari[2]:.2f}".replace('.', ',')) 
        self.form_layout.addRow("Ürün Adı:", self.ad_input)
        self.form_layout.addRow("Kategori:", self.kategori_input)
        self.form_layout.addRow("Fiyat (₺):", self.fiyat_input)
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.form_layout.addWidget(self.buttonBox)

    def accept(self):
        """Kullanıcı girişlerini doğrular ve kabul eder."""
        yeni_ad = self.ad_input.text().strip().capitalize()
        fiyat_str = self.fiyat_input.text().strip().replace(',', '.')
        
        if not yeni_ad:
            QMessageBox.warning(self, "Hata", "Ürün Adı boş olamaz.")
            return

        try:
            fiyat = float(fiyat_str)
            if fiyat < 0:
                 QMessageBox.warning(self, "Hata", "Fiyat sıfırdan küçük olamaz.")
                 return
            
            # Tüm kontroller başarılı, ebeveyn accept'i çağır
            super().accept()
        except ValueError:
            QMessageBox.warning(self, "Hata", "Lütfen geçerli bir fiyat (sayısal) girin.")
            # Hata olduğu için return etmiyoruz, kullanıcı düzeltmeli.

    def get_data(self):
        """Kabul edilen verileri döndürür."""
        # Fiyatı kaydetmeden önce yeniden float'a çeviriyoruz
        fiyat = float(self.fiyat_input.text().strip().replace(',', '.'))
        return (self.ad_input.text().strip().capitalize(), 
                self.kategori_input.text().strip(), 
                fiyat)


# --- Ana Stok Widget ---
class StokUygulamasiWidget(QWidget):
    """Ana stok listesi ve ürün ekleme/düzenleme işlevlerini içerir."""
    def __init__(self, veritabani_yoneticisi, status_bar):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.status_bar = status_bar
        self.arayuz_olustur()
        self.stogu_guncelle_arayuz()

    def arayuz_olustur(self):
        ana_duzen = QVBoxLayout(self)
        ust_duzen = QHBoxLayout()
        self.arama_input = QLineEdit()
        self.arama_input.setPlaceholderText("Ürün adında ara...")
        self.arama_input.textChanged.connect(self.tabloyu_filtrele)
        ust_duzen.addWidget(self.arama_input)
        self.yeni_urun_goster_btn = QPushButton(QIcon.fromTheme("list-add"), " Yeni Ürün Ekle")
        self.yeni_urun_goster_btn.clicked.connect(self.yeni_urun_formu_goster_gizle)
        ust_duzen.addWidget(self.yeni_urun_goster_btn)
        ana_duzen.addLayout(ust_duzen)

        # Yeni Ürün Ekleme Formu
        self.ekleme_formu_frame = QFrame()
        self.ekleme_formu_frame.setFrameShape(QFrame.StyledPanel)
        self.ekleme_formu_frame.setFrameShadow(QFrame.Raised)
        ekleme_duzen = QFormLayout(self.ekleme_formu_frame)
        self.yeni_urun_input = QLineEdit()
        self.yeni_kategori_input = QLineEdit()
        self.yeni_fiyat_input = QLineEdit()
        self.yeni_miktar_input = QLineEdit()
        
        self.onayla_ekle_btn = QPushButton(QIcon.fromTheme("dialog-ok"), " Onayla")
        self.onayla_ekle_btn.setDefault(True)
        
        ekleme_duzen.addRow("Ürün Adı:", self.yeni_urun_input)
        ekleme_duzen.addRow("Kategori:", self.yeni_kategori_input)
        ekleme_duzen.addRow("Fiyat (₺):", self.yeni_fiyat_input)
        ekleme_duzen.addRow("Miktar:", self.yeni_miktar_input)
        ekleme_duzen.addRow(self.onayla_ekle_btn)
        self.ekleme_formu_frame.hide()
        ana_duzen.addWidget(self.ekleme_formu_frame)
        self.onayla_ekle_btn.clicked.connect(self.yeni_urun_ekle)

        # Stok Tablosu
        self.stok_tablosu = QTableWidget()
        self.stok_tablosu.setColumnCount(5)
        self.stok_tablosu.setHorizontalHeaderLabels(["Ürün Adı", "Kategori", "Fiyat", "Miktar", "İşlemler"])
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.stok_tablosu.setEditTriggers(QTableWidget.NoEditTriggers) # Hücrelerin doğrudan düzenlenmesini engelle
        ana_duzen.addWidget(self.stok_tablosu)

    def stogu_guncelle_arayuz(self):
        """Veritabanından stok listesini çeker ve tabloyu günceller."""
        self.stok_tablosu.setRowCount(0)
        urun_listesi = self.veritabani.urunleri_getir()
        self.stok_tablosu.setRowCount(len(urun_listesi))
        for satir, (ad, kategori, fiyat, miktar) in enumerate(urun_listesi):
            # QTableWidgetItem'ı ayarla
            item_ad = QTableWidgetItem(ad)
            item_ad.setFlags(item_ad.flags() & ~Qt.ItemIsEditable) # Düzenlemeyi engelle
            
            self.stok_tablosu.setItem(satir, 0, item_ad)
            self.stok_tablosu.setItem(satir, 1, QTableWidgetItem(kategori))
            # Fiyat ve Miktar için sayısal gösterim (sıralama için iyileştirilebilir)
            self.stok_tablosu.setItem(satir, 2, QTableWidgetItem(f"{fiyat:.2f} ₺"))
            self.stok_tablosu.setItem(satir, 3, QTableWidgetItem(str(miktar)))
            
            # İşlem Butonu
            menu_btn = QPushButton("...")
            menu_btn.setFixedWidth(40)
            # Lambda içine urun_adi'nı (ad) ve butonu geçir
            menu_btn.clicked.connect(lambda ch, a=ad, b=menu_btn: self.guncelle_menusu_goster(a, b))
            self.stok_tablosu.setCellWidget(satir, 4, menu_btn)

    def guncelle_menusu_goster(self, urun_adi, buton):
        """Ürün için bağlam menüsünü gösterir."""
        menu = QMenu(self)
        
        duzenle_action = QAction(QIcon.fromTheme("document-edit"), "Bilgileri Düzenle", self)
        duzenle_action.triggered.connect(lambda: self.urun_duzenle(urun_adi))
        menu.addAction(duzenle_action)
        menu.addSeparator()
        
        artir_action = QAction(QIcon.fromTheme("go-up"), "Stok Artır/Ekle", self)
        artir_action.triggered.connect(lambda: self.miktar_girdi_goster(urun_adi, 'artır'))
        menu.addAction(artir_action)
        
        eksilt_action = QAction(QIcon.fromTheme("go-down"), "Stok Eksilt/Çıkar", self)
        eksilt_action.triggered.connect(lambda: self.miktar_girdi_goster(urun_adi, 'eksilt'))
        menu.addAction(eksilt_action)
        menu.addSeparator()
        
        sil_action = QAction(QIcon.fromTheme("edit-delete"), "Ürünü Sil", self)
        sil_action.triggered.connect(lambda: self.urun_sil(urun_adi))
        menu.addAction(sil_action)
        
        menu.exec_(buton.mapToGlobal(buton.rect().bottomLeft()))

    def urun_duzenle(self, urun_adi):
        """Seçilen ürünün detaylarını düzenleme diyalogu ile açar."""
        urun_detaylari = self.veritabani.urun_detay_getir(urun_adi)
        if not urun_detaylari:
            QMessageBox.warning(self, "Hata", "Ürün bulunamadı.")
            return
            
        dialog = UrunDuzenlemeDialog(urun_detaylari, self)
        
        # Sadece accept() ile onaylanan veriyi al
        if dialog.exec_() == QDialog.Accepted:
            yeni_ad, yeni_kategori, yeni_fiyat = dialog.get_data()
            self.veritabani.urun_detay_guncelle(urun_adi, yeni_ad, yeni_kategori, yeni_fiyat)
            self.stogu_guncelle_arayuz()
            self.status_bar.showMessage(f"'{yeni_ad}' güncellendi.", 3000)

    def yeni_urun_ekle(self):
        """Yeni ürün ekler veya mevcut ürünün miktarını artırır."""
        ad = self.yeni_urun_input.text().strip().capitalize()
        kat = self.yeni_kategori_input.text().strip()
        f_str = self.yeni_fiyat_input.text().strip().replace(',', '.')
        m_str = self.yeni_miktar_input.text().strip()
        
        if not all([ad, f_str, m_str]):
            QMessageBox.warning(self, "Uyarı", "Ürün Adı, Fiyat ve Miktar zorunludur.")
            return
            
        try:
            f = float(f_str)
            m = int(m_str)
            # Miktar pozitif olmalı, fiyat sıfır veya pozitif olabilir
            assert m > 0 and f >= 0 
        except (ValueError, AssertionError):
            QMessageBox.warning(self, "Hata", "Fiyat ve Miktar geçerli pozitif sayı olmalıdır.")
            return
            
        self.veritabani.urun_ekle(ad, kat, f, m)
        self.stogu_guncelle_arayuz()
        self.yeni_urun_formu_goster_gizle()
        self.status_bar.showMessage(f"'{ad}' eklendi/güncellendi.", 3000)
        
        # Alanları temizle
        for item in [self.yeni_urun_input, self.yeni_kategori_input, self.yeni_fiyat_input, self.yeni_miktar_input]:
            item.clear()

    def tabloyu_filtrele(self, metin):
        """Tablodaki ürünleri arama metnine göre filtreler."""
        for i in range(self.stok_tablosu.rowCount()):
            item = self.stok_tablosu.item(i, 0)
            if item:
                self.stok_tablosu.setRowHidden(i, metin.lower() not in item.text().lower())

    def yeni_urun_formu_goster_gizle(self):
        """Yeni ürün ekleme formunun görünürlüğünü değiştirir."""
        if self.ekleme_formu_frame.isVisible():
            self.ekleme_formu_frame.hide()
            self.yeni_urun_goster_btn.setText(" Yeni Ürün Ekle")
        else:
            self.ekleme_formu_frame.show()
            self.yeni_urun_input.setFocus()
            self.yeni_urun_goster_btn.setText(" Formu Kapat")


    def miktar_girdi_goster(self, u_adi, mod):
        """Stok artırma/eksiltme için miktar giriş diyalogu açar."""
        fiil = "Artır" if mod == 'artır' else "Eksilt"
        m, ok = QInputDialog.getInt(self, f"Stoğu {fiil}", f"'{u_adi}' için miktarı girin:", 1, 1, 999999)
        if ok:
            self.stok_miktari_degistir(u_adi, m if mod == 'artır' else -m)

    def stok_miktari_degistir(self, u_adi, m_farki):
        """Ürünün stok miktarını günceller."""
        mevcut = self.veritabani.mevcut_miktar_getir(u_adi)
        if mevcut + m_farki < 0:
            QMessageBox.warning(self, "Uyarı", "Stok eksiye düşemez.")
            return
        self.veritabani.urun_guncelle(u_adi, m_farki)
        self.stogu_guncelle_arayuz()
        self.status_bar.showMessage(f"'{u_adi}' stoğu güncellendi.", 3000)

    def urun_sil(self, u_adi):
        """Seçilen ürünü siler."""
        onay = QMessageBox.question(self, "Silmeyi Onayla", f"'{u_adi}' adlı ürünü silmek istediğinize emin misiniz?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if onay == QMessageBox.Yes:
            self.veritabani.urun_sil(u_adi)
            self.stogu_guncelle_arayuz()
            self.status_bar.showMessage(f"'{u_adi}' silindi.", 3000)


# --- Ana Pencere ---
class AnaPencere(QMainWindow):
    """Uygulamanın ana penceresi."""
    def __init__(self, kullanici_adi, veritabani_yoneticisi):
        super().__init__()
        self.setWindowTitle(f"Stok Yönetimi - [{kullanici_adi}]")
        self.setGeometry(100, 100, 900, 700)
        self.status_bar = self.statusBar()
        self.merkezi_widget = StokUygulamasiWidget(veritabani_yoneticisi, self.status_bar)
        self.setCentralWidget(self.merkezi_widget)


# --- İlk Kurulum Penceresi ---
class IlkKurulumPenceresi(QWidget):
    """Veritabanında kullanıcı yoksa yönetici hesabı oluşturur."""
    kurulum_tamamlandi = pyqtSignal(str) # Kullanıcı adını yayınlar

    def __init__(self, veritabani_yoneticisi):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.setWindowTitle("İlk Kurulum")
        self.setGeometry(400, 400, 400, 250)
        duzen = QVBoxLayout(self)
        
        baslik = QLabel("<h2>Yönetici Hesabı Oluştur</h2>")
        baslik.setAlignment(Qt.AlignCenter)
        duzen.addWidget(baslik)
        
        self.k_adi = QLineEdit()
        self.k_adi.setPlaceholderText("Kullanıcı Adı")
        self.sifre = QLineEdit()
        self.sifre.setPlaceholderText("Şifre")
        self.sifre.setEchoMode(QLineEdit.Password)
        self.sifre_t = QLineEdit()
        self.sifre_t.setPlaceholderText("Şifre Tekrar")
        self.sifre_t.setEchoMode(QLineEdit.Password)
        btn = QPushButton("Hesabı Oluştur")
        btn.setDefault(True)
        btn.clicked.connect(self.hesap_olustur)
        
        duzen.addWidget(self.k_adi)
        duzen.addWidget(self.sifre)
        duzen.addWidget(self.sifre_t)
        duzen.addWidget(btn)

    def hesap_olustur(self):
        k, s, st = self.k_adi.text().strip(), self.sifre.text(), self.sifre_t.text()
        if not k or not s:
            QMessageBox.warning(self, "Hata", "Kullanıcı Adı ve Şifre boş bırakılamaz.")
            return
        if s != st:
            QMessageBox.warning(self, "Hata", "Şifreler eşleşmiyor.")
            return
            
        b, m = self.veritabani.kullanici_ekle(k, s)
        if b:
            # Başarılı: Sinyali yay ve pencereyi kapat
            self.kurulum_tamamlandi.emit(k)
            self.close()
        else:
            QMessageBox.critical(self, "Hata", m)


# --- Kullanıcı Bilgileri Değiştir Penceresi ---
class KullaniciDegistirPenceresi(QWidget):
    """Kullanıcı adı ve şifre değiştirme penceresi."""
    # Değişiklik başarılı olduğunda Login penceresine dönmek için sinyal
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
        
        btn = QPushButton("Değişiklikleri Onayla")
        btn.setDefault(True)
        btn.clicked.connect(self.bilgileri_degistir)
        
        duzen.addStretch()
        duzen.addWidget(btn)

    def bilgileri_degistir(self):
        ek, es = self.e_kadi.text().strip(), self.e_sifre.text()
        yk, ys, yst = self.y_kadi.text().strip(), self.y_sifre.text(), self.y_sifre_t.text()
        
        if not all([ek, es, yk, ys, yst]):
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
            # Başarılı değişiklikten sonra sinyali yayınla
            self.degisiklik_yapildi.emit()
            self.close()
        else:
            QMessageBox.warning(self, "Hata", m)


# --- Giriş Penceresi ---
class GirisPenceresi(QWidget):
    """Kullanıcı girişini yöneten pencere."""
    login_basarili = pyqtSignal(str) # Kullanıcı adını yayınlar
    degistirme_penceresi_iste = pyqtSignal() # Bilgi değiştirme penceresi isteğini yayınlar

    def __init__(self, veritabani_yoneticisi):
        super().__init__()
        self.veritabani = veritabani_yoneticisi
        self.setWindowTitle("Giriş")
        self.setGeometry(400, 400, 400, 250)
        duzen = QVBoxLayout(self)
        
        baslik = QLabel("<h2>Giriş Yap</h2>")
        baslik.setAlignment(Qt.AlignCenter)
        duzen.addWidget(baslik)
        
        self.k_adi = QLineEdit()
        self.k_adi.setPlaceholderText("Kullanıcı Adı")
        self.sifre = QLineEdit()
        self.sifre.setPlaceholderText("Şifre")
        self.sifre.setEchoMode(QLineEdit.Password)
        
        login_btn = QPushButton("Giriş Yap")
        login_btn.setDefault(True)
        login_btn.clicked.connect(self.login_kontrol)
        
        degistir_btn = QPushButton("Bilgileri Değiştir")
        degistir_btn.setStyleSheet("border: none; color: #55aaff;")
        degistir_btn.clicked.connect(self.degistirme_penceresi_iste.emit)
        
        duzen.addWidget(self.k_adi)
        duzen.addWidget(self.sifre)
        duzen.addWidget(login_btn)
        duzen.addWidget(degistir_btn, 0, Qt.AlignRight)
        
        self.sifre.returnPressed.connect(self.login_kontrol)

    def login_kontrol(self):
        k, s = self.k_adi.text().strip(), self.sifre.text()
        
        if not k or not s:
            QMessageBox.warning(self, "Hata", "Lütfen Kullanıcı Adı ve Şifre girin.")
            return

        if self.veritabani.kullanici_dogrula(k, s):
            # Başarılı giriş: Önce sinyali yay, sonra pencereyi kapat
            self.login_basarili.emit(k)
            self.close()
        else:
            QMessageBox.warning(self, "Hata", "Kullanıcı adı veya şifre hatalı.")


# --- ANA KONTROLCU (DÜZELTİLMİŞ) ---
class AnaKontrolcu:
    """Uygulamanın başlangıcını, pencere geçişlerini ve akışını yönetir."""
    def __init__(self):
        self.veritabani = VeritabaniYoneticisi()
        # Referansları sakla ki pencere GC (Garbage Collection) ile yok olmasın
        self.mevcut_pencere = None 
        self.ana_pencere = None
        self.degistirme_penceresi = None

    def baslat(self):
        """Uygulamayı başlatır ve ilk pencereyi (Kurulum/Giriş) belirler."""
        if self.veritabani.kullanici_sayisi_getir() == 0:
            # Kullanıcı yoksa İlk Kurulum
            self.mevcut_pencere = IlkKurulumPenceresi(self.veritabani)
            self.mevcut_pencere.kurulum_tamamlandi.connect(self.ana_pencereyi_goster)
        else:
            # Kullanıcı varsa Giriş
            self.mevcut_pencere = GirisPenceresi(self.veritabani)
            self.mevcut_pencere.login_basarili.connect(self.ana_pencereyi_goster)
            self.mevcut_pencere.degistirme_penceresi_iste.connect(self.degistirme_penceresini_goster)
        
        self.mevcut_pencere.show()

    def ana_pencereyi_goster(self, kullanici_adi):
        """Ana stok yönetim penceresini açar."""
        self.ana_pencere = AnaPencere(kullanici_adi, self.veritabani)
        self.ana_pencere.show()
        
        # Önceki pencereyi (Giriş/Kurulum) kapat ve referansı temizle
        if self.mevcut_pencere:
            self.mevcut_pencere.close()
            self.mevcut_pencere = None

    def degistirme_penceresini_goster(self):
        """Kullanıcı Bilgileri Değiştirme penceresini açar."""
        # Giriş penceresi açıksa kapat
        if self.mevcut_pencere:
            self.mevcut_pencere.close()
            self.mevcut_pencere = None
            
        # Yeni değiştirme penceresini oluştur ve referansını sakla
        self.degistirme_penceresi = KullaniciDegistirPenceresi(self.veritabani)
        self.degistirme_penceresi.degisiklik_yapildi.connect(self.login_penceresine_don)
        self.degistirme_penceresi.show()

    def login_penceresine_don(self):
        """Değişiklikten sonra Giriş penceresini yeniden açar."""
        # Değiştirme penceresini kapat ve referansı temizle
        if self.degistirme_penceresi:
            self.degistirme_penceresi.close()
            self.degistirme_penceresi = None

        # Giriş penceresini YENİDEN oluştur ve referanslarını bağla
        # Bu, çökme riskini azaltan en güvenli yöntemdir.
        self.mevcut_pencere = GirisPenceresi(self.veritabani)
        self.mevcut_pencere.login_basarili.connect(self.ana_pencereyi_goster)
        self.mevcut_pencere.degistirme_penceresi_iste.connect(self.degistirme_penceresini_goster)
        self.mevcut_pencere.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Koyu Tema CSS stili
    app.setStyleSheet("""
        QWidget { background-color: #2b2b2b; color: #f0f0f0; font-family: Segoe UI; font-size: 10pt; }
        QMainWindow, QDialog { border: 1px solid #444; border-radius: 8px; }
        QLineEdit, QTableWidget, QDialog { background-color: #3c3c3c; border: 1px solid #555; border-radius: 4px; padding: 5px; }
        QPushButton { 
            background-color: #55aaff; color: #1e1e1e; border: none; padding: 8px 16px; border-radius: 4px; 
            font-weight: bold;
        }
        QPushButton:hover { background-color: #77bbff; }
        QPushButton:pressed { background-color: #3388ff; }
        
        QFrame { border: 1px solid #444; border-radius: 6px; padding: 10px; background-color: #323232; }

        QHeaderView::section { background-color: #444; padding: 6px; border: 1px solid #555; font-weight: bold;}
        QTableWidget { gridline-color: #555; }
        QTableWidget QTableCornerButton::section { background-color: #444; border: none; }

        QStatusBar { color: #aaa; background-color: #3c3c3c; }
        QMenu { background-color: #3c3c3c; border: 1px solid #555; border-radius: 4px;}
        QMenu::item { padding: 5px 15px; }
        QMenu::item:selected { background-color: #55aaff; color: #1e1e1e; }
        QMessageBox { background-color: #2b2b2b; }
    """)
    
    kontrolcu = AnaKontrolcu()
    kontrolcu.baslat()
    sys.exit(app.exec_())


