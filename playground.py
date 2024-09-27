from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
import requests
import os
import time
import pyautogui as gui
import img2pdf
import re
from PIL import Image
import io
import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QComboBox,
    QPushButton,
    QVBoxLayout, QMessageBox,
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt


def scrollDown(long, x, y):
    start_time = time.time()  # Waktu sebelum eksekusi

    gui.leftClick(x=x, y=y, duration=0.05)
    for i in range(long):
        print(f"scroll tick-{i}")
        gui.scroll(-10)

    end_time = time.time()  # Waktu setelah eksekusi

    print(f"Scrolling took {end_time - start_time} seconds")


class ImageToPDFConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.list_bmp = [
            "IDIK401303, Karil",
            "PAJA321002,Pengantar Ilmu Administrasi",
            "ADBI4531,Teori Pembuatan Keputusan",
            "ADPU444203,Sistem Informasi Manajemen (Edisi 3)",
            "IEEKMA4116,Management",
            "IEESPA4110,Introduction To Macroeconomics",
            "ISIP411002,Pengantar Sosiologi (Edisi 2)",
            "KIMD4110,Kimia Dasar I",
            "MATA411002,Kalkulus I (Edisi 2)",
            "MATA4303,Riset Operasi",
            "MATA4343,Riset Operasional I",
            "MATA434302,Riset Operasional I (Edisi 2)",
            "MKDU410902,Ilmu Sosial dan Budaya Dasar (Edisi 2)",
            "MKDU411103,Pendidikan Kewarganegaraan (Edisi 3)",
            "MKDU422102,Pendidikan Agama Islam (Edisi 2)",
            "MKWI4201,Bahasa Inggris",
            "MKWU410202,Pendidikan Agama Katolik (Edisi 2)",
            "MKWU4103,Pendidikan Agama Kristen",
            "MKWU4104,Pendidikan Agama Buddha",
            "MKWU4105,Pendidikan Agama Hindu",
            "MKWU4107,Pendidikan Agama Khonghucu",
            "MKWU410802,Bahasa Indonesia (Edisi 2)",
            "SATS4111,Komputer 1",
            "SATS4120,Matematika 1",
            "SATS412002,Matematika I (Edisi 2)",
            "SATS412102,Metode Statistik 1 (Edisi 2)",
            "SATS4122,Aljabar Linear Terapan",
            "SATS4210,Matematika 2",
            "SATS421002,Matematika II (Edisi 2)",
            "SATS4211,Metode Statistik 2",
            "SATS4212,Analisis Data Statistik",
            "SATS4213,Pengumpulan dan Penyajian Data",
            "SATS422002,Matematika III (Edisi 2)",
            "SATS4221,Pengantar Probabilitas",
            "SATS422302,Komputer 2",
            "SATS4224,Pengantar Sosiometri",
            "SATS4312,Model Linear Terapan",
            "SATS4313,Demografi",
            "SATS431302,Demografi (Edisi 2)",
            "SATS432102,Metode Sampling (Edisi 2)",
            "SATS4322,Pengantar Proses Stokastik",
            "SATS432302,Metode Peramalan (Edisi 2)",
            "SATS432303,Metode Peramalan (Edisi 3)",
            "SATS4324,Inferensi Bayesian",
            "SATS4410,Pengantar Statistika Matematis 1",
            "SATS4411,Metode Statistika Non Parametrik",
            "SATS4420,Pengantar Statistika Matematis 2",
            "SATS442102,Metode Statistika Multivariat (Edisi 2)",
            "SATS4422,Metode Sekuensial",
            "SATS4423,Analisis Runtun Waktu",
            "MSIM4312,Metodelogi Penelitian (Edisi 2)",
            "STAT4215,Statistika Pengawasan Kualitas",
            "STAT4231,Pengumpulan dan Penyajian Data",
            "STAT4331,Asuransi I",
            "STAT433102,Asuransi I (Edisi 2)",
            "STAT4410,Pengantar Proses Stokastik I",
            "STAT4431,Rancangan Percobaan",
            "STAT4434,Asuransi II",
            "STAT4435,Pengantar Sosiometri",
        ]
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Image to PDF Converter")

        # Create a font for larger text
        large_font = QFont()
        large_font.setPointSize(16)  # Adjust the font size as needed

        self.label_bmp = QLabel("BMP:")
        self.label_bmp.setFont(large_font)  # Apply the larger font

        # Create a combo box for BMP selection
        self.combo_bmp = QComboBox()
        list_bmp = [
            "IDIK401303, Karil",
            "PAJA321002,Pengantar Ilmu Administrasi",
            "ADBI4531,Teori Pembuatan Keputusan",
            "ADPU444203,Sistem Informasi Manajemen (Edisi 3)",
            "IEEKMA4116,Management",
            "IEESPA4110,Introduction To Macroeconomics",
            "ISIP411002,Pengantar Sosiologi (Edisi 2)",
            "KIMD4110,Kimia Dasar I",
            "MATA411002,Kalkulus I (Edisi 2)",
            "MATA4303,Riset Operasi",
            "MATA4343,Riset Operasional I",
            "MATA434302,Riset Operasional I (Edisi 2)",
            "MKDU410902,Ilmu Sosial dan Budaya Dasar (Edisi 2)",
            "MKDU411103,Pendidikan Kewarganegaraan (Edisi 3)",
            "MKDU422102,Pendidikan Agama Islam (Edisi 2)",
            "MKWI4201,Bahasa Inggris",
            "MKWU410202,Pendidikan Agama Katolik (Edisi 2)",
            "MKWU4103,Pendidikan Agama Kristen",
            "MKWU4104,Pendidikan Agama Buddha",
            "MKWU4105,Pendidikan Agama Hindu",
            "MKWU4107,Pendidikan Agama Khonghucu",
            "MKWU410802,Bahasa Indonesia (Edisi 2)",
            "SATS4111,Komputer 1",
            "SATS4120,Matematika 1",
            "SATS412002,Matematika I (Edisi 2)",
            "SATS412102,Metode Statistik 1 (Edisi 2)",
            "SATS4122,Aljabar Linear Terapan",
            "SATS4210,Matematika 2",
            "SATS421002,Matematika II (Edisi 2)",
            "SATS4211,Metode Statistik 2",
            "SATS4212,Analisis Data Statistik",
            "SATS4213,Pengumpulan dan Penyajian Data",
            "SATS422002,Matematika III (Edisi 2)",
            "SATS4221,Pengantar Probabilitas",
            "SATS422302,Komputer 2",
            "SATS4224,Pengantar Sosiometri",
            "SATS4312,Model Linear Terapan",
            "SATS4313,Demografi",
            "SATS431302,Demografi (Edisi 2)",
            "SATS432102,Metode Sampling (Edisi 2)",
            "SATS4322,Pengantar Proses Stokastik",
            "SATS432302,Metode Peramalan (Edisi 2)",
            "SATS432303,Metode Peramalan (Edisi 3)",
            "SATS4324,Inferensi Bayesian",
            "SATS4410,Pengantar Statistika Matematis 1",
            "SATS4411,Metode Statistika Non Parametrik",
            "SATS4420,Pengantar Statistika Matematis 2",
            "SATS442102,Metode Statistika Multivariat (Edisi 2)",
            "SATS4422,Metode Sekuensial",
            "SATS4423,Analisis Runtun Waktu",
            "MSIM431202,Metodelogi Penelitian (Edisi 2)",
            "STAT4215,Statistika Pengawasan Kualitas",
            "STAT4231,Pengumpulan dan Penyajian Data",
            "STAT4331,Asuransi I",
            "STAT433102,Asuransi I (Edisi 2)",
            "STAT4410,Pengantar Proses Stokastik I",
            "STAT4431,Rancangan Percobaan",
            "STAT4434,Asuransi II",
            "STAT4435,Pengantar Sosiometri",
        ]

        self.combo_bmp.addItems(sorted([item.split(",")[1] for item in list_bmp]))



        self.label_modul = QLabel("Modul:")
        self.label_modul.setFont(large_font)  # Apply the larger font

        # Create a combo box for Modul selection
        self.combo_modul = QComboBox()
        self.combo_modul.addItems([str(i) for i in range(1, 15)])

        self.convert_button = QPushButton("Convert")
        # Connect the button to the run_script method without passing any argument
        self.convert_button.clicked.connect(self.run_script)
        self.status_label = QLabel("")
        self.status_label.setFont(large_font)

        layout = QVBoxLayout()
        layout.addWidget(self.label_bmp)
        layout.addWidget(self.combo_bmp)
        layout.addWidget(self.label_modul)
        layout.addWidget(self.combo_modul)
        layout.addWidget(self.convert_button)
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def run_script(self):
        self.convert_button.setEnabled(False)  # Disable the button
        try:
            bmp_description = self.combo_bmp.currentText()
            bmp_code = None
            for item in self.list_bmp:
                if item.endswith(bmp_description):
                    bmp_code = item.split(",")[0]
                    bmp_name = item.split(",")[1]
                    break

            if bmp_code is None:
                self.status_label.setText("Error: BMP not found in the list")
                return

            modul = self.combo_modul.currentText()
            main(bmp_code, bmp_name, int(modul))
            self.status_label.setText(f"Conversion completed!\n{bmp_code}-{bmp_name}")
            print(f"Conversion completed for\n{bmp_code}-{bmp_name} =================>")
        except Exception as e:
            # Print the error to the output console
            sys.excepthook(type(e), e, e.__traceback__)
        finally:
            # Enable the button again whether there was an error or not
            self.convert_button.setEnabled(True)


# Redirect errors to the output console
def exception_hook(exctype, value, traceback):
    sys.__excepthook__(exctype, value, traceback)


sys.excepthook = exception_hook


def main(bmp, bmp_name, modul):
    print("Buat folder untuk menyimpan gambar jika belum ada")
    folder_images = os.path.join("data", "images")
    folder_pdf = os.path.join("data", "modul", bmp_name)
    # module_name = bmp.split(",")[1].strip()  # new
    # folder_pdf = os.path.join(folder_pdf, module_name)  # new
    os.makedirs(folder_pdf, exist_ok=True)  # new
    if not os.path.exists(folder_images):
        os.makedirs(folder_images)
    else:
        files = os.listdir(folder_images)
        if files:
            for file in files:
                os.remove(os.path.join(folder_images, file))

    # URL dasar
    base_url = f"https://pustaka.ut.ac.id/reader/services/view.php?doc=M{modul}&format=jpg&subfolder={bmp}/&page="

    # For Chrome:
    # options = webdriver.ChromeOptions()
    # driver = webdriver.Chrome(options=options)

    # For Mozilla:
    options = webdriver.FirefoxOptions()
    driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
    # driver = webdriver.Firefox(options=options)

    print("Buka halaman login")
    driver.get(f"https://pustaka.ut.ac.id/reader/index.php?modul={bmp}")

    print('Tunggu sampai elemen "username" dan "password" muncul')
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )

    print("Masukkan username dan password")
    username_input.send_keys("mahasiswa")
    password_input.send_keys("utpeduli")

    # Dapatkan seluruh teks dalam elemen form
    form_text = (
        WebDriverWait(driver, 10)
        .until(EC.presence_of_element_located((By.ID, "form1")))
        .text
    )

    # Cari teks captcha dalam teks form
    captcha_question = re.search(r"Berapa hasil dari \d+ \+ \d+ =", form_text).group()

    # Ekstrak angka dari pertanyaan captcha
    numbers = re.findall(r"\d+", captcha_question)
    print(numbers)

    # Hitung hasil penjumlahan
    captcha_answer = sum(int(number) for number in numbers)
    print(captcha_answer)

    # Masukkan hasil penjumlahan ke dalam elemen input
    captcha_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "ccaptcha"))
    )
    captcha_input.send_keys(str(captcha_answer))

    # Temukan tombol submit
    submit_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "submit"))
    )

    # Klik tombol submit
    submit_button.click()

    time.sleep(2)

    # Tunggu sampai elemen dapat diinteraksi, lalu klik
    modul1_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//a[normalize-space()='[Modul {modul}]']")
        )
    )
    modul1_link.click()

    time.sleep(10)

    # Tunggu sampai elemen dengan XPath "//img[@title='Single Page']" muncul, lalu klik
    single_page_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//img[@title='Single Page']"))
    )
    single_page_button.click()

    print("Tunggu beberapa detik untuk memastikan login berhasil")
    time.sleep(5)

    print("Headers yang sama dengan browser")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "Referer": base_url,
    }

    # Assuming driver is your WebDriver instance
    # element_hal_terakhir = driver.find_element_by_xpath(
    #     "//div[@class='flowpaper_lblTotalPages flowpaper_tblabel flowpaper_numberOfPages']"
    # )
    element_hal_terakhir = driver.find_element(
        By.CSS_SELECTOR,
        ".flowpaper_lblTotalPages.flowpaper_tblabel.flowpaper_numberOfPages",
    )
    print(f"element hal terakhir : {element_hal_terakhir}")
    text_hal_terakhir = (
        element_hal_terakhir.text
    )  # This should give you something like " / 38"

    # Use a regular expression to extract the number
    match = re.search(r"\d+", text_hal_terakhir)
    if match:
        hal_terakhir = int(match.group())  # Convert the string to an integer


    msg = QMessageBox()
    msg.setIcon(QMessageBox.Question)
    msg.setText("Do you want to continue?")
    msg.setInformativeText("This will download all the images from the website.")
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    msg.setDefaultButton(QMessageBox.No)
    ret = msg.exec_()
    if ret == QMessageBox.Yes:
        pass
    else:
        sys.exit(0)





    print("Looping untuk mendownload gambar")
    for i in range(1, hal_terakhir + 1):  # Ganti 5 dengan jumlah halaman + 1
        url = base_url + str(i)
        driver.get(url)

        print(f"Dapatkan URL gambar-{i} aktual dari halaman yang dibuka oleh Selenium")
        img_url = (
            WebDriverWait(driver, 10)
            .until(EC.presence_of_element_located((By.TAG_NAME, "img")))
            .get_attribute("src")
        )

        print(f"Dapatkan cookie-{i} dari Selenium")
        cookies = driver.get_cookies()

        print(f"Buat session-{i} requests dan tambahkan cookie dari Selenium")
        s = requests.Session()
        for cookie in cookies:
            s.cookies.set(cookie["name"], cookie["value"])
        s.headers.update(headers)

        print(f"Gunakan session-{i} untuk mengunduh gambar")
        response = s.get(img_url)

        print(f"Pastikan respons-{i} adalah sukses")
        if response.status_code == 200:
            print(f"Menulis gambar-{i} ke file")
            with open(os.path.join(folder_images, f"image_{i}.jpeg"), "wb") as f:
                f.write(response.content)
        else:
            print(f"Failed to download page {i}")

    # Tutup browser
    driver.quit()

    # Daftar semua file gambar
    image_files = [
        os.path.join(folder_images, f"image_{i}.jpeg")
        for i in range(1, hal_terakhir + 1)
    ]

    # List untuk menyimpan BytesIO objects dari setiap gambar
    image_bytes = []

    # Baca setiap file gambar sebagai bytes dan tambahkan ke list
    for image_file in image_files:
        with open(image_file, "rb") as f:
            img = Image.open(f)
            byte_arr = io.BytesIO()
            img.save(byte_arr, format="JPEG")
            image_bytes.append(byte_arr.getvalue())
    pdf_filename = os.path.join(folder_pdf, f"{bmp} - modul {modul}.pdf")
    # Konversi list dari BytesIO objects menjadi PDF
    with open(pdf_filename, "wb") as f:
        f.write(img2pdf.convert(image_bytes))
    files = os.listdir(folder_images)
    if files:
        for file in files:
            os.remove(os.path.join(folder_images, file))


if __name__ == "__main__":
    # main(bmp="SATS442102", modul=5)
    app = QApplication(sys.argv)
    window = ImageToPDFConverter()
    window.show()
    sys.exit(app.exec_())

