from NesineUserInfo import usernema,password
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import xlsxwriter
import time

class Nesine:
    def __init__(self,usernema,password):
        self.browser = webdriver.Chrome()
        self.usernema = usernema
        self.password = password
        self.lstmacLig = []
        self.lstmacTarih= []
        self.lstmacSaat = []
        self.lstmacAdi = []
        self.lstmacLink = []
    def singIn(self):
        url = "https://www.nesine.com/iddaa"
        self.browser.get(url) 
        
        #Kullanıcı Adı ve Şifre Girişi Yapıp Enter Tuşuna Basıldı
        self.browser.find_element(By.ID,"txtUsername").send_keys(self.usernema)
        self.browser.find_element(By.ID,"realpass").send_keys(self.password)
        self.browser.find_element(By.ID,"realpass").send_keys(Keys.ENTER)
        
        #Açılan Tarayıcı Çerezleri Kabul Etme
        self.browser.find_element(By.ID,"c-p-bn").send_keys(Keys.ENTER)

        #Açılan Tarayıcı Sayfa Sonuna Alma
        self.browser.execute_script("window.scrollBy(0,document.body.scrollHeight)")
        time.sleep(5)
    def Bülten(self):
        sections = self.browser.find_elements(By.TAG_NAME,"section")
        for s in sections:
            macLİgTarih = s.find_elements(By.CLASS_NAME,"name-date-col")
            for mlt in macLİgTarih:                
                maclar = s.find_elements(By.CSS_SELECTOR,".odd-col.event-list.pre-event")
                for m in maclar:
                    self.lstmacLig.append(mlt.find_element(By.CLASS_NAME,"name").text)
                    self.lstmacTarih.append(mlt.find_element(By.CLASS_NAME,"date").text)
                    self.lstmacSaat.append(m.find_element(By.CLASS_NAME,"code-time-name").find_element(By.CLASS_NAME,"time").text)
                    self.lstmacAdi.append(m.find_element(By.CLASS_NAME,"code-time-name").find_element(By.CLASS_NAME,"name").text)
                    self.lstmacLink.append(m.find_element(By.CLASS_NAME,"code-time-name").find_element(By.CLASS_NAME,"name").find_element(By.TAG_NAME,"a").get_attribute("href"))
    def Excel_Aktar(self):
        #excel dosyasının adını belirleyelim
        workbook = xlsxwriter.Workbook( "Bülten.xlsx" )  
        #çalışma sayfası ekleyelim 
        worksheet = workbook.add_worksheet("Data")
        #worksheet.write(satir,0,veri) A=0 sütunu sabittir
        for satir,veri in enumerate(self.lstmacLig):
            worksheet.write(satir,0,veri)
        for satir,veri in enumerate(self.lstmacTarih):
            worksheet.write(satir,1,veri)
        for satir,veri in enumerate(self.lstmacSaat):
            worksheet.write(satir,2,veri)
        for satir,veri in enumerate(self.lstmacAdi):
            worksheet.write(satir,3,veri)
        for satir,veri in enumerate(self.lstmacLink):
            worksheet.write(satir,4,veri)
        #dosya işlemlerini bitiriyoruz.
        workbook.close()
nesine = Nesine(usernema,password)
nesine.singIn()
nesine.Bülten()
nesine.Excel_Aktar()
        
