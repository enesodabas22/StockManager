import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QMessageBox,
                             QTableWidget, QTableWidgetItem, QMenu, QAction,
                             QInputDialog, QHeaderView)


stok = {
    "Elma": 50,
    "Muz": 30,
    "Portakal": 45,
    "Çilek": 100
}


class StokUygulamasi(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 Stok Yönetimi (Özel Miktar Güncelleme)")
        self.setGeometry(100, 100, 750, 500)

        self.arayuz_olustur()
        self.stogu_guncelle_arayuz()

    def arayuz_olustur(self):
        ana_duzen = QVBoxLayout()


        ekle_duzen = QHBoxLayout()
        ekle_duzen.addWidget(QLabel("Yeni Ürün Adı:"))
        self.yeni_urun_input = QLineEdit()
        ekle_duzen.addWidget(self.yeni_urun_input)

        ekle_duzen.addWidget(QLabel("İlk Miktar:"))
        self.yeni_miktar_input = QLineEdit()
        ekle_duzen.addWidget(self.yeni_miktar_input)

        self.ekle_btn = QPushButton("Yeni Ürün Ekle")
        self.ekle_btn.clicked.connect(self.yeni_urun_ekle)
        ekle_duzen.addWidget(self.ekle_btn)


        self.stok_tablosu = QTableWidget()
        self.stok_tablosu.setColumnCount(3)
        self.stok_tablosu.setHorizontalHeaderLabels(["Ürün Adı", "Miktar", "İşlem"])


        self.stok_tablosu.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # Miktar sütunu genişlesin
        self.stok_tablosu.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)


        ana_duzen.addLayout(ekle_duzen)
        ana_duzen.addWidget(QLabel("\n--- Güncel Stok Durumu ---"))
        ana_duzen.addWidget(self.stok_tablosu)

        self.setLayout(ana_duzen)



    def stogu_guncelle_arayuz(self):


        self.stok_tablosu.setRowCount(0)
        sirali_stok = sorted(stok.items())
        self.stok_tablosu.setRowCount(len(sirali_stok))

        satir = 0
        for urun, miktar in sirali_stok:

            self.stok_tablosu.setItem(satir, 0, QTableWidgetItem(urun))


            self.stok_tablosu.setItem(satir, 1, QTableWidgetItem(str(miktar)))


            menu_btn = QPushButton("...")
            menu_btn.setStyleSheet("width: 30px;")


            menu_btn.clicked.connect(lambda checked, row=satir: self.guncelle_menusu_goster(row, menu_btn))

            self.stok_tablosu.setCellWidget(satir, 2, menu_btn)

            satir += 1

        self.stok_tablosu.resizeColumnsToContents()

    def guncelle_menusu_goster(self, satir_indeksi, buton):
        """Üç nokta butonuna tıklandığında bağlam menüsünü gösterir."""

        urun_adi = self.stok_tablosu.item(satir_indeksi, 0).text()

        menu = QMenu(self)

        # Stok Artır/Ekle Eylemi
        artir_action = QAction("Miktar Artır/Ekle", self)
        # Fonksiyonu 'artir' moduyla çağırır
        artir_action.triggered.connect(lambda: self.miktar_girdi_goster(urun_adi, 'artır'))
        menu.addAction(artir_action)

        # Stok Eksilt/Çıkar Eylemi
        eksilt_action = QAction("Miktar Eksilt/Çıkar", self)
        # Fonksiyonu 'eksilt' moduyla çağırır
        eksilt_action.triggered.connect(lambda: self.miktar_girdi_goster(urun_adi, 'eksilt'))
        menu.addAction(eksilt_action)

        # Menüyü butonun hemen altında göster
        menu.exec_(buton.mapToGlobal(buton.rect().bottomRight()))

    def miktar_girdi_goster(self, urun_adi, mod):
        """Kullanıcıdan artırma/eksiltme miktarını girmesini isteyen pencereyi açar."""

        # QInputDialog kullanarak sayı girmesini isteyelim
        dialog = QInputDialog()
        dialog.setWindowTitle(f"{urun_adi} Stok {mod.capitalize()}")
        dialog.setLabelText(f"Lütfen {urun_adi} için {mod}ülecek miktarı girin (Pozitif Sayı):")
        dialog.setInputMode(QInputDialog.IntInput)
        dialog.setIntRange(1, 999999)  # 1'den başlasın
        dialog.setIntValue(1)  # Başlangıç değeri 1 olsun

        ok = dialog.exec_()
        miktar_str = dialog.intValue()  # QInputDialog'un int inputu string döner

        if ok:
            try:
                miktar = int(miktar_str)
            except ValueError:
                QMessageBox.critical(self, "Hata", "Geçersiz miktar. Lütfen sadece tam sayı girin.")
                return

            if mod == 'eksilt':
                # Eksiltme işlemi için miktar negatif olmalı
                miktar *= -1

            self.stok_miktari_degistir(urun_adi, miktar)

    def stok_miktari_degistir(self, urun_adi, miktar_farki):
        """Stok miktarını verilen fark kadar artırır veya eksiltir."""
        global stok

        if urun_adi in stok:
            yeni_miktar = stok[urun_adi] + miktar_farki

            if yeni_miktar < 0:
                QMessageBox.warning(self, "Uyarı",
                                    f"'{urun_adi}' stoğu sıfırın altına düşürülemez. Mevcut: {stok[urun_adi]}")
                return

            stok[urun_adi] = yeni_miktar

            if miktar_farki > 0:
                mesaj = f"'{urun_adi}' stoğuna **{miktar_farki}** adet eklendi. Yeni stok: {yeni_miktar}"
            else:
                mesaj = f"'{urun_adi}' stoğundan **{-miktar_farki}** adet düşüldü. Yeni stok: {yeni_miktar}"

            self.stogu_guncelle_arayuz()
            QMessageBox.information(self, "Güncellendi", mesaj)
        else:
            QMessageBox.critical(self, "Hata", "Ürün stokta bulunamadı.")

    def yeni_urun_ekle(self):
        """Arayüzün üstündeki input'lardan yeni bir ürün ekler."""
        urun_adi = self.yeni_urun_input.text().strip().capitalize()
        miktar_str = self.yeni_miktar_input.text().strip()

        # ... (Giriş kontrolleri ve ekleme mantığı aynı kaldı)
        if not urun_adi or not miktar_str:
            QMessageBox.warning(self, "Uyarı", "Lütfen ürün adı ve ilk miktarını giriniz.")
            return

        try:
            miktar = int(miktar_str)
            if miktar <= 0:
                QMessageBox.warning(self, "Uyarı", "Miktar pozitif bir sayı olmalıdır.")
                return
        except ValueError:
            QMessageBox.warning(self, "Hata", "Miktar için geçerli bir sayı giriniz.")
            return

        if urun_adi in stok:
            stok[urun_adi] += miktar
        else:
            stok[urun_adi] = miktar

        self.stogu_guncelle_arayuz()
        self.yeni_urun_input.clear()
        self.yeni_miktar_input.clear()
        QMessageBox.information(self, "Başarılı", f"'{urun_adi}' ürünü güncellendi/eklendi.")


# --- Uygulamayı Başlatma ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    pencere = StokUygulamasi()
    pencere.show()
    sys.exit(app.exec_())