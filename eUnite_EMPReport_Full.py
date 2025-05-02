from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import os
import time
from datetime import datetime

class EUniteBot:
    def __init__(self):
        # Step 0: กำหนด download directory
        self.download_dir = r"C:\Users\weerawat.m\Downloads\EMP"
        os.makedirs(self.download_dir, exist_ok=True)
        
        options = Options()
        options.binary_location = os.path.join(os.path.dirname(__file__), "Chrome", "chrome-win64", "chrome.exe")
        options.add_experimental_option('prefs', {
            'download.default_directory': self.download_dir,
            'download.prompt_for_download': False,
            'safebrowsing.enabled': True
        })
        options.add_argument('--start-maximized')
        options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
        
        service = Service(os.path.join(os.path.dirname(__file__), "chromedriver-win64", "chromedriver.exe"))
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 10)
        self.actions = ActionChains(self.driver)

    def login(self, username, password):
        try:
            # Step 1: เข้า URL
            self.driver.get("https://signin.eunite.com/eunite/login?account=sbg")
            time.sleep(2)
            
            # Step 2: Log in
            username_field = self.wait.until(EC.presence_of_element_located((By.NAME, "username")))
            username_field.clear()
            username_field.send_keys(username)
            
            password_field = self.wait.until(EC.presence_of_element_located((By.NAME, "password")))
            password_field.clear()
            password_field.send_keys(password + Keys.RETURN)
            time.sleep(2)
            
            # Step 3: เปลี่ยนสิทธิ์เป็น Administrator
            dropdown = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-toggle='dropdown']")))
            self.driver.execute_script("arguments[0].click();", dropdown)
            time.sleep(1)
            
            admin = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Administrator']")))
            self.driver.execute_script("arguments[0].click();", admin)
            time.sleep(2)
            
            return True
        except Exception as e:
            print(f"Login failed: {e}")
            return False

    def download_report(self):
        try:
            # Step 4: เข้า URL รายงาน
            self.driver.get("https://apd.eunite.com/eunite/record-2014R3/RecordList.do?method=form&rec=RM_EmployeeXLS_All&")
            time.sleep(2)
            
            # Step 5: ใส่วันที่ปัจจุบัน
            formatted_date = datetime.now().strftime('%d%m') + str(datetime.now().year + 543)
            print(f"กรอกวันที่: {formatted_date}")
            self.driver.execute_script("document.elementFromPoint(588, 183).click();")
            self.actions.send_keys(formatted_date + Keys.RETURN).perform()
            time.sleep(1)
            
            # Step 6: กด Tab 10 ครั้งและเลือก Main
            print("กด Tab 10 ครั้งและเลือก Main")
            [self.actions.send_keys(Keys.TAB).perform() for _ in range(10)]
            self.wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Main']"))).click()
            time.sleep(1)
            
            # Step 7: กด Tab 3 ครั้ง, เลือก XLSX และระบบจะเริ่มดาวน์โหลดอัตโนมัติ
            print("กด Tab 3 ครั้ง และเลือก XLSX เพื่อเริ่มดาวน์โหลด")
            [self.actions.send_keys(Keys.TAB).perform() for _ in range(3)]
            self.actions.send_keys(Keys.SPACE, Keys.ARROW_DOWN, Keys.RETURN).perform()
            time.sleep(2)
            
            # Step 8: รอดาวน์โหลดและเปลี่ยนชื่อไฟล์
            download_success = self._wait_for_download(formatted_date)
            if download_success:
                print("การดาวน์โหลดเสร็จสมบูรณ์")
                return True
            return False
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการดาวน์โหลด: {e}")
            return False

    def _wait_for_download(self, new_filename, timeout=60):
        start = time.time()
        print(f"รอการดาวน์โหลดไฟล์... (timeout {timeout} วินาที)")
        
        # เก็บรายชื่อไฟล์เดิมและแสดงรายการ
        existing_files = [f for f in os.listdir(self.download_dir) if f.endswith('.xlsx')]
        print(f"จำนวนไฟล์ก่อนดาวน์โหลด: {len(existing_files)} ไฟล์")
        print(f"รายการไฟล์เดิม: {', '.join(sorted(existing_files))}")
        existing_file_set = set(os.path.join(self.download_dir, f) for f in existing_files)
        
        while time.time() - start < timeout:
            try:
                current_files = os.listdir(self.download_dir)
                
                # ตรวจสอบว่ามีไฟล์ที่กำลังดาวน์โหลดอยู่หรือไม่
                if not any(f.endswith('.crdownload') for f in current_files):
                    # ค้นหาไฟล์ใหม่ที่เพิ่งดาวน์โหลดเสร็จ
                    current_xlsx_files = set(os.path.join(self.download_dir, f) 
                                          for f in current_files if f.endswith('.xlsx'))
                    new_files = current_xlsx_files - existing_file_set
                    
                    if new_files:  # ถ้าพบไฟล์ใหม่
                        latest_file = max(new_files, key=os.path.getctime)
                        
                        if os.path.getsize(latest_file) > 0:
                            # กำหนดชื่อไฟล์ใหม่
                            counter = 1
                            new_target = os.path.join(self.download_dir, f"{new_filename}.xlsx")
                            while os.path.exists(new_target):
                                new_target = os.path.join(self.download_dir, f"{new_filename}-{counter}.xlsx")
                                counter += 1
                            
                            # เปลี่ยนชื่อไฟล์ใหม่
                            os.rename(latest_file, new_target)
                            print(f"ดาวน์โหลดเสร็จสิ้น และบันทึกเป็น: {os.path.basename(new_target)}")
                            
                            # แสดงรายการไฟล์หลังดาวน์โหลด
                            final_files = [f for f in os.listdir(self.download_dir) if f.endswith('.xlsx')]
                            print(f"จำนวนไฟล์หลังดาวน์โหลด: {len(final_files)} ไฟล์")
                            print(f"รายการไฟล์ทั้งหมด: {', '.join(sorted(final_files))}")
                            return True
                        else:
                            print("ไฟล์ที่ดาวน์โหลดมีขนาดเป็น 0")
                            return False
            except Exception as e:
                print(f"เกิดข้อผิดพลาดในการตรวจสอบดาวน์โหลด: {e}")
            time.sleep(2)
        
        print("หมดเวลารอดาวน์โหลด")
        return False

    def close(self):
        try:
            if hasattr(self, 'driver'):
                print("Step 9: ปิด Chrome")
                self.driver.quit()
        except Exception as e:
            print(f"เกิดข้อผิดพลาดในการปิด Chrome: {e}")

def main():
    bot = EUniteBot()
    try:
        if bot.login("weerawat.m@somboon.co.th", "WwMm@123456789"):
            print("เข้าสู่ระบบสำเร็จ")
            if bot.download_report():
                print("ดาวน์โหลดรายงานสำเร็จ")
            else:
                print("ดาวน์โหลดล้มเหลว")
        else:
            print("เข้าสู่ระบบล้มเหลว")
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")
    finally:
        bot.close()

if __name__ == "__main__":
    main()
