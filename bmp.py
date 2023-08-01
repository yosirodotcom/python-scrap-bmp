from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import os
import time
import pyautogui as gui
import img2pdf
import re
from PIL import Image
import io

# the last code

print("Buat folder untuk menyimpan gambar jika belum ada")
folder_images = os.path.join("MyCoding", "BMPUT", "images")
folder_pdf = os.path.join("MyCoding", "BMPUT", "modul")

os.makedirs(folder_images, exist_ok=True)
os.makedirs(folder_pdf, exist_ok=True)


# URL dasar
base_url = "https://pustaka.ut.ac.id/reader/services/view.php?doc=M1&format=jpg&subfolder=SATS4411/&page="

print("Inisialisasi webdriver")
driver = webdriver.Chrome()

print("Buka halaman login")
driver.get("https://pustaka.ut.ac.id/reader/index.php?modul=SATS4411")

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
    EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='[Modul 1]']"))
)
modul1_link.click()

time.sleep(5)

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

print("Scroll down to load all images")


def scrollDown(long):
    start_time = time.time()  # Waktu sebelum eksekusi

    gui.leftClick(x=1221, y=456, duration=0.05)
    for i in range(long):
        gui.scroll(-50)

    end_time = time.time()  # Waktu setelah eksekusi

    print(f"Scrolling took {end_time - start_time} seconds")


scrollDown(long=(90 * 2))


print("Looping untuk mendownload gambar")
for i in range(1, 5):  # Ganti 5 dengan jumlah halaman + 1
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
image_files = [os.path.join(folder_images, f"image_{i}.jpeg") for i in range(1, 10)]

# List untuk menyimpan BytesIO objects dari setiap gambar
image_bytes = []

# Baca setiap file gambar sebagai bytes dan tambahkan ke list
for image_file in image_files:
    with open(image_file, "rb") as f:
        img = Image.open(f)
        byte_arr = io.BytesIO()
        img.save(byte_arr, format="JPEG")
        image_bytes.append(byte_arr.getvalue())
pdf_filename = os.path.join(folder_pdf, "output.pdf")
# Konversi list dari BytesIO objects menjadi PDF
with open(pdf_filename, "wb") as f:
    f.write(img2pdf.convert(image_bytes))
