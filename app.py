from flask import Flask, render_template, request, jsonify, send_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import os
import re
import time
import json

app = Flask(__name__)

# تكوين المجلدات
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'templates')
if not os.path.exists(TEMPLATES_DIR):
    os.makedirs(TEMPLATES_DIR)

# متغير عام لتخزين جلسة المتصفح
driver = None
session_file = "whatsapp_session"

def get_driver():
    """إنشاء أو إعادة استخدام جلسة المتصفح"""
    global driver
    if driver is None:
        chrome_options = Options()
        chrome_options.add_argument(f"--user-data-dir={session_file}")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.get("https://web.whatsapp.com")
            
            # انتظار تسجيل الدخول لأول مرة
            wait = WebDriverWait(driver, 120)
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']")))
            print("✅ تم تسجيل الدخول بنجاح")
        except Exception as e:
            print(f"❌ خطأ في تشغيل المتصفح: {e}")
            driver = None
            raise
    
    return driver

def extract_members(group_input):
    """استخراج أعضاء الجروب"""
    driver = get_driver()
    wait = WebDriverWait(driver, 20)
    
    try:
        # إذا كان المدخل رابطاً
        if "chat.whatsapp.com" in group_input:
            driver.get(group_input)
            time.sleep(5)
            
            # البحث عن زر الانضمام إذا كان رابط دعوة
            try:
                join_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='_ak8q']")))
                join_button.click()
                time.sleep(3)
            except:
                pass  # ربما هو بالفعل عضو في الجروب
        
        # البحث عن الجروب
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']")))
        search_box.clear()
        search_box.send_keys(group_input)
        time.sleep(2)
        
        # النقر على الجروب من نتائج البحث
        try:
            group_chat = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@title='{group_input}']")))
            group_chat.click()
        except:
            # محاولة بديلة
            group_chat = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, '_ak8q')]")))
            group_chat.click()
        
        time.sleep(2)
        
        # فتح معلومات الجروب
        try:
            # أيقونة المعلومات (ثلاث نقاط أو اسم الجروب)
            group_header = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-testid='chat-header']")))
            group_header.click()
            time.sleep(1)
            
            # النقر على "معلومات الجروب"
            info_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='معلومات الجروب']")))
            info_option.click()
        except:
            # طريقة بديلة
            group_name_elem = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='_ak8q']")))
            group_name_elem.click()
        
        time.sleep(3)
        
        # استخراج الأعضاء
        members = []
        
        # التمرير لأسفل لتحميل كل الأعضاء
        members_panel = driver.find_element(By.XPATH, "//div[@data-testid='group-participants']")
        last_height = 0
        while True:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", members_panel)
            time.sleep(1)
            new_height = driver.execute_script("return arguments[0].scrollHeight", members_panel)
            if new_height == last_height:
                break
            last_height = new_height
        
        # استخراج أسماء وأرقام الأعضاء
        member_elements = driver.find_elements(By.XPATH, "//div[@data-testid='cell-frame-container']")
        
        for elem in member_elements:
            try:
                # محاولة استخراج الاسم
                name_element = elem.find_element(By.XPATH, ".//span[contains(@class, '_ao3e')]")
                name = name_element.text.strip()
                
                # محاولة استخراج الرقم
                phone = ""
                try:
                    phone_element = elem.find_element(By.XPATH, ".//div[contains(@class, '_akbu')]")
                    phone_text = phone_element.text
                    # استخراج الأرقام من النص
                    phone_numbers = re.findall(r'\+?\d[\d\s\-\(\)]{8,}\d', phone_text)
                    if phone_numbers:
                        phone = re.sub(r'[\s\-\(\)]', '', phone_numbers[0])
                except:
                    pass
                
                if name and name != "أنت":
                    members.append({
                        "الاسم": name,
                        "رقم الهاتف": phone
                    })
            except:
                continue
        
        if not members:
            # محاولة طريقة أخرى
            member_elements = driver.find_elements(By.XPATH, "//div[contains(@class, '_ak8q')]//span")
            for elem in member_elements:
                text = elem.text.strip()
                if text and text != "أنت" and len(text) > 1:
                    phone = re.search(r'\+?\d{10,15}', text)
                    members.append({
                        "الاسم": text if not phone else text.replace(phone.group(), "").strip(),
                        "رقم الهاتف": phone.group() if phone else ""
                    })
        
        # حفظ في ملف Excel
        if members:
            df = pd.DataFrame(members)
            df.to_excel("members.xlsx", index=False, engine='openpyxl')
            return len(members), members
        else:
            raise Exception("لم يتم العثور على أعضاء. تأكد من أنك في الجروب الصحيح")
            
    except TimeoutException:
        raise Exception("انتهى الوقت. تأكد من اتصال الإنترنت ووجود الجروب")
    except Exception as e:
        raise Exception(f"حدث خطأ: {str(e)}")

@app.route('/')
def index():
    """الصفحة الرئيسية"""
    return render_template('index.html')

@app.route('/export', methods=['POST'])
def export_members():
    """استخراج الأعضاء"""
    group_input = request.form.get('group_url', '').strip()
    
    if not group_input:
        return jsonify({'success': False, 'message': 'يرجى إدخال رابط الجروب أو اسمه'})
    
    try:
        count, members = extract_members(group_input)
        return jsonify({
            'success': True, 
            'message': f'✅ تم استخراج {count} عضو بنجاح',
            'count': count
        })
    except Exception as e:
        return jsonify({'success': False, 'message': f'❌ {str(e)}'})

@app.route('/download', methods=['GET'])
def download_file():
    """تحميل ملف الإكسيل"""
    file_path = "members.xlsx"
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name="whatsapp_members.xlsx")
    return jsonify({'success': False, 'message': 'لا يوجد ملف للتحميل'}), 404

@app.route('/check_file', methods=['GET'])
def check_file():
    """التحقق من وجود ملف"""
    return jsonify({'exists': os.path.exists("members.xlsx")})

@app.route('/clear_session', methods=['POST'])
def clear_session():
    """مسح جلسة واتساب (للاختبار)"""
    global driver
    if driver:
        driver.quit()
        driver = None
    return jsonify({'success': True, 'message': 'تم مسح الجلسة'})

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
