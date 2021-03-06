from flask import Flask,render_template, request,jsonify,redirect
from werkzeug import secure_filename
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options 
import pyautogui
import pandas as pd
import webbrowser
import shutil
import os

ALLOWED_EXTENSIONS = set(['csv'])

app = Flask(__name__)

webbrowser.open('http://127.0.0.1:5000/', new=2)

@app.route("/")
def index():
    return render_template('home.html')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS    

def delete_files():
    directory = os.path.dirname(os.path.realpath(__file__)) + "\\static\\upload"
    print(directory)
    for the_file in os.listdir(directory):
        file_path = os.path.join(directory, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception as e:
            print(e)    


@app.route("/excel_upload", methods = ['GET', 'POST'])    
def excel_upload():
    delete_files()
    f = request.files['file']
    if f and allowed_file(f.filename):
        f.save(os.path.join("static/upload","input.csv"))
    return redirect('/')

@app.route("/automate")
def scrape():
    mode = request.args.get('mode', 0, type=int)
    email = request.args.get('email', 0, type=str)
    password = request.args.get('password', 0, type=str)
    interior = request.args.get('interior', 0, type=str)
    cover_url = request.args.get('cover', 0, type=str)

    chrome_options = Options()
#    chrome_options.add_argument("--headless") 
    prefs = {'profile.managed_default_content_settings.images':2}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_path = r"chromedriver.exe" 
    driver = webdriver.Chrome(chrome_path,chrome_options=chrome_options)
    driver.maximize_window()
    #driver.set_window_position(-10000,-1000)
    url = "https://kdp.amazon.com/en_US/"
    driver.get(url)
    try:
        element_present = EC.presence_of_element_located((By.ID,"signinButton-announce"))
        WebDriverWait(driver, 100).until(element_present)
        driver.find_element_by_id("signinButton-announce").click()
        time.sleep(1)
    except TimeoutException:
        print("timeout!")

    try:
        element_present = EC.presence_of_element_located((By.ID,"ap_email"))
        WebDriverWait(driver, 100).until(element_present)
        driver.find_element_by_id("ap_email").send_keys(email)
        driver.find_element_by_id("ap_password").send_keys(password)
        driver.find_element_by_id("signInSubmit").click()
        time.sleep(1)
    except TimeoutException:
        print("timeout!")
    
    csv_file = "static\\upload\\input.csv"
    if not os.path.exists(csv_file):
        res = {}
        res['res'] = "empty"
        return jsonify(res)
    
    if mode == 0:
        headers = ["title","author","description","keyword1","keyword2","keyword3","keyword4","keyword5","keyword6","keyword7","upload","cover","trim_size"]
        df = pd.read_csv(csv_file,skiprows=1,names=headers)
        titles = df.title.tolist()
        authors = df.author.tolist()
        descriptions = df.description.tolist()
        keyword1 = df.keyword1.tolist()
        keyword2 = df.keyword2.tolist()
        keyword3 = df.keyword3.tolist()
        keyword4 = df.keyword4.tolist()
        keyword5 = df.keyword5.tolist()
        keyword6 = df.keyword6.tolist()
        keyword7 = df.keyword7.tolist()
        upload = df.upload.tolist()
        cover = df.cover.tolist()
        book_size = df.trim_size.tolist()

        for i in range(0,len(titles)):
            print(titles[i])
            try:
                element_present = EC.presence_of_element_located((By.ID,"create-paperback-button"))
                WebDriverWait(driver, 100).until(element_present)
                driver.find_element_by_id("create-paperback-button").click()
                time.sleep(1)
            except TimeoutException:
                print("timeout!")
            
            try:
                element_present = EC.presence_of_element_located((By.ID,"data-print-book-title"))
                WebDriverWait(driver, 20).until(element_present)

                driver.find_element_by_id("data-print-book-title").clear()
                driver.find_element_by_id("data-print-book-title").send_keys(titles[i])
                driver.find_element_by_id("data-print-book-primary-author-last-name").clear()
                driver.find_element_by_id("data-print-book-primary-author-last-name").send_keys(authors[i])
                driver.find_element_by_id("data-print-book-description").clear()
                driver.find_element_by_id("data-print-book-description").send_keys(descriptions[i])
                driver.find_element_by_id("non-public-domain").click()
                driver.find_element_by_id("data-print-book-keywords-0").clear()
                driver.find_element_by_id("data-print-book-keywords-0").send_keys(keyword1[i])
                driver.find_element_by_id("data-print-book-keywords-1").clear()
                driver.find_element_by_id("data-print-book-keywords-1").send_keys(keyword2[i])
                driver.find_element_by_id("data-print-book-keywords-2").clear()
                driver.find_element_by_id("data-print-book-keywords-2").send_keys(keyword3[i])
                driver.find_element_by_id("data-print-book-keywords-3").clear()
                driver.find_element_by_id("data-print-book-keywords-3").send_keys(keyword4[i])
                driver.find_element_by_id("data-print-book-keywords-4").clear()
                driver.find_element_by_id("data-print-book-keywords-4").send_keys(keyword5[i])
                driver.find_element_by_id("data-print-book-keywords-5").clear()
                driver.find_element_by_id("data-print-book-keywords-5").send_keys(keyword6[i])
                driver.find_element_by_id("data-print-book-keywords-6").clear()
                driver.find_element_by_id("data-print-book-keywords-6").send_keys(keyword7[i])
                driver.find_element_by_id("data-print-book-categories-button-proto-announce").click()
                time.sleep(1)
                driver.find_element_by_id("checkbox-non--classifiable").click()
                driver.find_element_by_xpath("""//*[@id="category-chooser-ok-button"]/span/input""").click()
                time.sleep(0.5)
                driver.find_element_by_xpath("""//*[@id="data-print-book-large-print"]""").click()
                driver.find_element_by_xpath("""//*[@id="data-print-book-large-print"]""").click()

                # driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                # driver.find_element_by_xpath("""//*[@id="data-print-book-is-adult-content"]/div/div/div[1]/div/label/input""").click()

                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('space')
                time.sleep(1)
                driver.find_element_by_id("save-and-continue-announce").click()
                time.sleep(1)
            except TimeoutException as err:
                print(err)
            

            try:
                element_present = EC.presence_of_element_located((By.ID,"free-print-isbn-btn-announce"))
                WebDriverWait(driver, 10).until(element_present)
                driver.find_element_by_id("free-print-isbn-btn-announce").click()
                time.sleep(1)
            except TimeoutException:
                print("timeout!")

            try:
                element_present = EC.presence_of_element_located((By.ID,"print-isbn-confirm-button-announce"))
                WebDriverWait(driver, 10).until(element_present)
                driver.find_element_by_id("print-isbn-confirm-button-announce").click()
                time.sleep(1)
            except TimeoutException:
                print("timeout!")
            try:
                element_present = EC.presence_of_element_located((By.ID,"trim-size-btn-announce"))
                WebDriverWait(driver, 10).until(element_present)
                driver.find_element_by_id("trim-size-btn-announce").click()
                time.sleep(1)
            except TimeoutException:
                print("timeout!")
            print(book_size[i])
            
            filter_size = book_size[i]
            filter_size=filter_size.strip()

            
            try:
                element_present = EC.presence_of_element_located((By.ID,"trim-size-standard-option-1-3-announce"))
                WebDriverWait(driver, 10).until(element_present)
                template = driver.find_element_by_id("templateArea")
                popularbt0 = template.find_element_by_id("trim-size-popular-option-0-0-announce")
                populartemp0 = popularbt0.get_attribute("data-primary-display")
                popularbt1 = template.find_element_by_id("trim-size-popular-option-0-1-announce")
                populartemp1 = popularbt1.get_attribute("data-primary-display")
                popularbt2 = template.find_element_by_id("trim-size-popular-option-0-2-announce")
                populartemp2 = popularbt2.get_attribute("data-primary-display")
                popularbt3 = template.find_element_by_id("trim-size-popular-option-0-3-announce")
                populartemp3 = popularbt3.get_attribute("data-primary-display")
                standbt0 = template.find_element_by_id("trim-size-standard-option-0-0-announce")
                standtemp0 = standbt0.get_attribute("data-primary-display")
                standbt1 = template.find_element_by_id("trim-size-standard-option-0-1-announce")
                standtemp1 = standbt1.get_attribute("data-primary-display")
                standbt2 = template.find_element_by_id("trim-size-standard-option-0-2-announce")
                standtemp2 = standbt2.get_attribute("data-primary-display")
                standbt3 = template.find_element_by_id("trim-size-standard-option-0-3-announce")
                standtemp3 = standbt3.get_attribute("data-primary-display")
                standbt4 = template.find_element_by_id("trim-size-standard-option-1-0-announce")
                standtemp4 = standbt4.get_attribute("data-primary-display")
                standbt5 = template.find_element_by_id("trim-size-standard-option-1-1-announce")
                standtemp5 = standbt5.get_attribute("data-primary-display")
                standbt6 = template.find_element_by_id("trim-size-standard-option-1-2-announce")
                standtemp6 = standbt6.get_attribute("data-primary-display")
                standbt7 = template.find_element_by_id("trim-size-standard-option-1-3-announce")
                standtemp7 = standbt7.get_attribute("data-primary-display")
                nonbt0 = template.find_element_by_id("trim-size-nonstandard-option-0-0-announce")
                nontemp0 = nonbt0.get_attribute("data-primary-display")
                nonbt1 = template.find_element_by_id("trim-size-nonstandard-option-0-1-announce")
                nontemp1 = nonbt1.get_attribute("data-primary-display")
                nonbt2 = template.find_element_by_id("trim-size-nonstandard-option-0-2-announce")
                nontemp2 = nonbt2.get_attribute("data-primary-display")
                nonbt3 = template.find_element_by_id("trim-size-nonstandard-option-0-3-announce")
                nontemp3 = nonbt3.get_attribute("data-primary-display")

                if populartemp0 == filter_size:
                    popularbt0.click()
                elif populartemp1 == filter_size:
                    popularbt1.click()
                elif populartemp2 == filter_size:
                    popularbt2.click()
                elif populartemp3 == filter_size:
                    popularbt3.click()
                elif standtemp0 == filter_size:
                    standbt0.click()
                elif standtemp1 == filter_size:
                    standbt1.click()
                elif standtemp2 == filter_size:
                    standbt2.click()
                elif standtemp3 == filter_size:
                    standbt3.click()
                elif standtemp4 == filter_size:
                    standbt4.click()
                elif standtemp5 == filter_size:
                    standbt5.click()
                elif standtemp6 == filter_size:
                    standbt6.click()
                elif standtemp7 == filter_size:
                    standbt7.click()
                elif nontemp0 == filter_size:
                    nonbt0.click()
                elif nontemp1 == filter_size:
                    nonbt1.click()
                elif nontemp2 == filter_size:
                    nonbt2.click()
                elif nontemp3 == filter_size:
                    nonbt3.click()

                time.sleep(1)
            except TimeoutException:
                print("timeout!")

            try:
                element_present = EC.presence_of_element_located((By.ID,"a-autoid-4-announce"))
                WebDriverWait(driver, 10).until(element_present)
                driver.find_element_by_id("a-autoid-4-announce").click()
                driver.find_element_by_id("a-autoid-6-announce").click()
            except TimeoutException:
                print("timeout!")

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight*3/5);")

            driver.find_element_by_id("data-print-book-publisher-interior-file-upload-browse-button-announce").click()
            upload_path = interior + "\\" + upload[i]
            counter = 0
            pyautogui.typewrite(interior[0])
            time.sleep(0.5)
            while counter < len(upload_path):
                pyautogui.typewrite(upload_path[counter])
                counter+= 1

            pyautogui.press('enter')


            flag = 0
            while flag == 0:
                try:
                    modal_style = driver.find_element_by_id("a-popover-lgtbox").get_attribute("style")
                    if modal_style.find("none") != -1:
                        flag = 1
                except Exception:
                    time.sleep(1)

            time.sleep(1)
            driver.find_element_by_xpath("""//*[@id="data-print-book-publisher-cover-choice-accordion"]/div[2]/div/div[1]/a""").click()
            time.sleep(2)
            driver.find_element_by_id("data-print-book-publisher-cover-file-upload-browse-button-announce").click()

            pyautogui.typewrite(cover_url[0])
            time.sleep(0.5)
            upload_path = cover_url + "\\" + cover[i]
            for ar in upload_path:
                print(ar)
                pyautogui.typewrite(ar)

            pyautogui.press('enter')

            flag = 0
            while flag == 0:
                try:
                    modal_style = driver.find_element_by_id("a-popover-lgtbox").get_attribute("style")
                    if modal_style.find("none") != -1:
                        flag = 1
                except Exception:
                    time.sleep(1)

            time.sleep(1)
            driver.find_element_by_id("save-announce").click()

            flag = 0
            while flag == 0:
                try:
                    modal_style = driver.find_element_by_id("a-popover-lgtbox").get_attribute("style")
                    if modal_style.find("none") != -1:
                        flag = 1
                except Exception:
                    time.sleep(1)


            driver.find_element_by_xpath("""//*[@id="top-0"]/div/div[2]/div/div[2]/span/a[1]""").click()
    else:
        headers = ["price"]
        df = pd.read_csv(csv_file,skiprows=1,names=headers)
        prices = df.price.tolist()
        
        try:
            element_present = EC.presence_of_element_located((By.ID,"podbookshelftable_filter_input"))
            WebDriverWait(driver, 200).until(element_present)
        except TimeoutException:
            print("timeout!")
            

        number = 1
        counter = -1
        driver.find_element_by_id("podbookshelftable_filter_input").click()
        driver.find_element_by_id("podbookshelftable-publishing-status-filter-draft").click()
        time.sleep(3)
        while number > 0:
            counter += 1
            
            table = driver.find_element_by_xpath("""//*[@id="podbookshelftable"]/div[4]/div/table""")
            trs = table.find_elements_by_class_name("mt-row")
            trs[0].find_element_by_class_name("indie-split-button-main-action-normal").click()
            
            try:
                element_present = EC.presence_of_element_located((By.ID,"book-setup-navigation-bar-content-link"))
                WebDriverWait(driver, 20).until(element_present)
                driver.find_element_by_id("book-setup-navigation-bar-content-link").click()
                time.sleep(1)
            except TimeoutException:
                driver.find_element_by_xpath("""//*[@id="top-0"]/div/div[2]/div/div[2]/span/a[1]""").click()
                continue
                print("timeout!")

            try:
                element_present = EC.presence_of_element_located((By.ID,"data-print-book-publication-date-input"))
                WebDriverWait(driver, 10).until(element_present)
                time.sleep(1)
            except TimeoutException:
                driver.find_element_by_xpath("""//*[@id="top-0"]/div/div[2]/div/div[2]/span/a[1]""").click()
                continue
                print("timeout!")
            accept_url = driver.current_url + "?acceptProof=CONVERTED"
            driver.get(accept_url)
            time.sleep(3)
            driver.find_element_by_id("save-and-continue-announce").click()

            try:
                element_present = EC.presence_of_element_located((By.ID,"print-book-worldwide-rights-field"))
                WebDriverWait(driver, 100).until(element_present)
                driver.find_element_by_id("print-book-worldwide-rights-field").click()
                driver.find_element_by_xpath("""//*[@id="data-pricing-print-us-price-input"]/input""").send_keys(str(prices[counter]))
                time.sleep(3)
            except TimeoutException:
                driver.find_element_by_xpath("""//*[@id="top-0"]/div/div[2]/div/div[2]/span/a[1]""").click()
                continue
                print("timeout!")
            driver.find_element_by_id("save-and-publish-announce").click()
            flag = 0
            while flag == 0:
                try:
                    modal_style = driver.find_element_by_id("a-popover-lgtbox").get_attribute("style")
                    if modal_style.find("none") != -1:
                        flag = 1
                except Exception:
                    time.sleep(1)
                
            try:
                element_present = EC.presence_of_element_located((By.ID,"publish-confirm-popover-print-close"))
                WebDriverWait(driver, 10).until(element_present)
                time.sleep(1)
                driver.find_element_by_id("publish-confirm-popover-print-close").click()
            except TimeoutException:
                driver.find_element_by_xpath("""//*[@id="top-0"]/div/div[2]/div/div[2]/span/a[1]""").click()
                continue
                print("timeout!")
            time.sleep(1)
                
            driver.find_element_by_id("podbookshelftable_filter_input").click()
            driver.find_element_by_id("podbookshelftable-publishing-status-filter-draft").click()
            time.sleep(3)

            table = driver.find_element_by_xpath("""//*[@id="podbookshelftable"]/div[4]/div/table""")
            trs = table.find_elements_by_class_name("mt-row")
            number = len(trs)


    res = {}
    res['res'] ="success" 
    return jsonify(res)

@app.route('/downloadfile/<path>')
def download(path=None):
    print(path)
    try:
        return send_file("/static/out/"+path, as_attachment=True) 
       #I also tried flask.send_file and send_static_file
    except:
        return("Generic error message")

if __name__=='__main__':
    app.run(debug = True)

