import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os

from datetime import date, datetime, timedelta
#Virtual Display
from pyvirtualdisplay import Display

# Fungsi untuk mengirim pesan dan menghitung jumlah pesan terkirim
def kirim_pesan(number, name, memiliki_wa, tidak_punya_wa):
    try:
        # display = Display(visible=0, size=(800, 600))
        # display.start()
        # Tentukan jalur ke ChromeDriver Anda
        chrome_driver_path = 'path_driver'

        # Inisialisasi Service object
        service = Service(chrome_driver_path)

        # Inisialisasi Options object
        options = Options()
        options.add_argument("user-data-dir=")

        # Inisialisasi WebDriver dengan menggunakan Service object dan Options object
        driver = webdriver.Chrome(service=service, options=options)

        url = 'https://api.whatsapp.com/send/?phone=62{}'.format(number)
        driver.get(url)
        time.sleep(2)

        # Klik lanjut chat
        wait = WebDriverWait(driver, 30)
        lanjut_path = "//a[contains(@title,\"Bagikan di WhatsApp\")]"
        lanjut = wait.until(EC.presence_of_element_located((By.XPATH, lanjut_path)))
        lanjut.click()
        time.sleep(2)

        # Gunakan WhatsApp Web
        gunakan_path = "//h4[@class='_9vd5'][2]/a[@class='_9vcv _9vcx']/span[@class='_advp _aeam']"
        gunakan = wait.until(EC.presence_of_element_located((By.XPATH, gunakan_path)))
        time.sleep(2)
        gunakan.click()

        # Masukkan pesan per paragraf
        message_box_path = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]'
        message_box = wait.until(EC.presence_of_element_located((By.XPATH, message_box_path)))
        message_box.click()
        # message_box.send_keys(pesan)
 
        # display.stop()
        memiliki_wa.append([number, name])  # Menambahkan nomor dan nama yang memiliki WhatsApp ke dalam list

    except TimeoutException:
        print("Terjadi kesalahan saat mengirim pesan kepada {}".format(number))
        print ("=============================================")
        tidak_punya_wa.append([number, name])  # Menambahkan nomor dan nama yang tidak memiliki WhatsApp ke dalam list


# Baca data phone.xlsx
read_phone = pd.read_excel('phone.xlsx')

# Inisialisasi variabel jumlah pesan terkirim dan jumlah tidak memiliki WhatsApp Web
memiliki_wa = []
tidak_punya_wa = []

# Kirim pesan untuk setiap nomor telepon
for i in range(len(read_phone)):
    number = str(read_phone['phone'][i])  # number berisi lebih dari 1 value
    name = str(read_phone['nama'][i])

    # Kirim pesan dan periksa apakah terkirim
    kirim_pesan(number, name, memiliki_wa, tidak_punya_wa)

# Simpan nomor dan nama yang memiliki WhatsApp ke dalam file Excel
df_memiliki_wa = pd.DataFrame(memiliki_wa, columns=['phone', 'nama'])
df_memiliki_wa.to_excel(datetime.now().strftime("%d-%m-%Y") + 'memiliki_wa.xlsx', index=False)

# Simpan nomor dan nama yang tidak memiliki WhatsApp ke dalam file Excel
df_tidak_punya_wa = pd.DataFrame(tidak_punya_wa, columns=['phone', 'nama'])
df_tidak_punya_wa.to_excel(datetime.now().strftime("%d-%m-%Y") + 'tidak_punya_wa.xlsx', index=False)

print ("Jumlah Number yang di check = {}".format(len(read_phone)))
print ("=================================================")
print ("Jumlah Memiliki WA = {}".format(len(memiliki_wa)))
print ("Jumlah Tidak Memiliki WA = {}".format(len(tidak_punya_wa)))
