import os
import sys
import glob
import shutil
import pymssql
import pandas as pd
from time import sleep
from random import randint
from datetime import datetime

# smpt protocol and email send libs
import mimetypes
import smtplib,ssl
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.multipart import MIMEMultipart

# selenium modules
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
 

URL = "https://www.bank.com"
LOGIN = "XXXXXXX"
PASSWORD = "XXXXX"

DW_SQL_CONN = {'DBIP':'0.0.0.0',
        'DBLOGIN' : 'login',
        'DBPASSWORD' : 'password'}


def get_exe_directory():
    if getattr(sys, 'frozen', False):
        # Running as a bundled executable
        return os.path.dirname(sys.executable)
    else:
        # Running as a script
        return os.path.dirname(os.path.abspath(__file__))

 
def open_website(url):

        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
        "download.default_directory": get_exe_directory(),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        })      

        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("--timeout=120000")   
        chrome_options.add_argument("--set-script-timeout=120000")
        chrome_options.add_argument("--disable-webgl")
        chrome_options.add_argument("--disable-extensions")
        
        servico = Service(ChromeDriverManager().install())

        driver = webdriver.Chrome(service=servico, options=chrome_options)
        driver.set_window_size(1300, 800)

        driver.get(url)

        return driver

 
def run_download_bills_RPA(LOGIN, PASSWORD, DATA_CONSULTA=datetime.now()):

        driver = open_website(URL)

        # Wait for and click the access button
        access_button = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="open_modal_more_access"]/div')))
        sleep(3)
        access_button.click()
        
        # Wait for and locate the modal field
        modal_field = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'idl-modal-more-access-body')))
        sleep(3)

        login_options = modal_field.find_element(by=By.XPATH, value= '//*[@id="idl-more-access-select-login"]')
        sleep(2)
        login_options.send_keys('operator code')

        # Wait for and click the operator modal field
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="idl-more-access-container-operator"]/div')))
        operator_modal_field = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idl-more-access-container-operator"]/div')))
        sleep(2)
        operator_modal_field.click()
        sleep(2)
        
        # Enter operator code
        LOGIN_field = operator_modal_field.find_element(by=By.XPATH, value='//*[@id="idl-more-access-input-operator"]')
        sleep(2)
        LOGIN_field.send_keys(LOGIN)
        sleep(2)

        # Access the website
        accer_website = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="idl-more-access-submit-button"]')))
        sleep(2)
        accer_website.click()

        # Enter password using the virtual keyboard
        keys = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/section/div[1]/div[2]/div/div[1]')))
        keys = keys.find_elements(by=By.TAG_NAME, value='a')
        for digit in PASSWORD:
                for key in keys:
                        if digit in key.text.split(' '):
                                sleep(0.8)
                                key.click()

        # Click the 'access' button
        key_access = [key.find_element(by=By.XPATH, value='//*[@id="acessar"]') for key in keys if 'access' in key.text.split(' ')][0]
        sleep(1.5)
        key_access.click()

        sleep(10)

        # Select options list
        options_list = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, f'[value="B"]')))
        options_list.click()

        # Continue to the next page
        continue_ = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="btn-habitilar-biometria-facial"]')))
        sleep(randint(1, 3))
        continue_.click()

        try:
                # Handle potential popup
                sleep(10)
                popup = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'dialog.modalPushNotification__container')))
                understood = popup.find_element(by=By.CSS_SELECTOR, value="button.Bank-button-custom")
                sleep(1)
                understood.click()
                sleep(2)
        except Exception as _:
                print(_)

        sleep(5)
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/header/div[6]/div/div/nav/ul/li[3]/a')))
        sleep(3)
        
        # Hover over the field option
        field_ = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/header/div[6]/div/div/nav/ul/li[3]/a')))
        sleep(3)
        a = ActionChains(driver)
        a.move_to_element(field_).perform()
        
        # Click the field option button
        field_option_button = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/header/div[6]/div/div/nav/ul/li[3]/ul/li[7]/a')))
        sleep(2)
        field_option_button.click()
        sleep(4)
        driver.execute_script("window.scrollTo(0, 450);")

        try:
                # Close potential popup
                sleep(10)
                popup = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ajuda-movel-angular-app"]/ajuda-movel/div/app-float-empresas/div')))
                close_button = popup.find_element(by=By.CLASS_NAME, value="voxel-icon__m")
                sleep(1)
                close_button.click()
                sleep(3)
        except Exception as _:
                pass

        # Download billings
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[5]/div[2]/div[2]/a')))
        billings_to_download = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[5]/div[2]/div[2]/a')))
        sleep(2)
        billings_to_download.click()
        sleep(10)
        driver.execute_script("window.scrollTo(0, 450);")
        sleep(3)

        screen_percent = 75
        script_scroll = f"window.scrollTo(0, (document.body.scrollHeight * {screen_percent / 100}));"

        # Input dates for statement request
        input_initial_date = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="voxel-input-1"]')))
        input_final_date = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="voxel-input-2"]')))
        input_initial_date.send_keys(DATA_CONSULTA)
        input_final_date.send_keys(DATA_CONSULTA)
        sleep(2)

        # Request statement
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, 'buscaMovimentacaoConsultaExtrato')))
        ask_statement = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, 'buscaMovimentacaoConsultaExtrato')))
        ask_statement.click()
        sleep(10)
        ask_statement.click()
        sleep(4)

        # Wait for statement to be ready
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/section/div/div/cobranca-francesinha/main/app-extract-transition-titles/section/div/div/div/div/section/div/section[2]/section[2]/app-movement-summarized/section[4]/app-extract-request/section[2]/voxel-alert[3]/div/div[2]')))
        sleep(2)
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/section/div/div/cobranca-francesinha/main/app-extract-transition-titles/section/div/div/div/div/section/div/section[2]/section[2]/app-movement-summarized/section[4]/app-extract-request/section[2]/voxel-alert[3]/div/div[2]/section/a')))
        search = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/div/div/cobranca-francesinha/main/app-extract-transition-titles/section/div/div/div/div/section/div/section[2]/section[2]/app-movement-summarized/section[4]/app-extract-request/section[2]/voxel-alert[3]/div/div[2]/section/a')))
        sleep(2)
        search.click()

        # Wait for and click the Excel download button
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="render-angular-app"]/cobranca-francesinha/main/app-query-extract-transition-titles/div/section/div/section[2]/app-table-request-history/table/tbody/tr[1]/td[5]/div/app-download-excel/button/voxel-icon')))
        sleep(1)
        excel_download = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="render-angular-app"]/cobranca-francesinha/main/app-query-extract-transition-titles/div/section/div/section[2]/app-table-request-history/table/tbody/tr[1]/td[5]/div/app-download-excel/button/voxel-icon')))
        sleep(1)
        driver.refresh()

        screen_percent = 2
        script_scroll = f"window.scrollTo(0, (document.body.scrollHeight * {screen_percent / 100}));"
        driver.execute_script(script_scroll)
        sleep(7)
        excel_download.click()

        # Rename and move the downloaded file
        script_directory = get_exe_directory()

        file_type = '*.xlsx'  

        file_path = glob.glob(os.path.join(script_directory, file_type))

        try:

            if len(file_path) == 1:
                excel_file = file_path[0]
                    
        except IndexError:
            pass

        FILE_DATE = str(DATA_CONSULTA).replace('/','_')

        new_name = f'TODAYS_BILLS_{FILE_DATE}.xlsx'    

        new_path = os.path.join(get_exe_directory() , new_name) 


        os.rename(excel_file, new_path)

        subfolder = os.path.join(script_directory, "PLANILHAS DO DIA") 

        if not os.path.exists(subfolder): 
            os.makedirs(subfolder)
        
        shutil.move(new_path, os.path.join(subfolder, new_name)) 


def get_file_path(current_directory = get_exe_directory()):

    file_type = '*.xlsx' 

    file_paths = glob.glob(os.path.join(current_directory, file_type))

    try:
        if len(file_paths) == 1:
            excel_path = file_paths[0]
            return excel_path
        else:
            pass    
    except Exception as e:
        print(e)


def verify_autenticy_in_database(file_path):

    path = file_path

    columns = ['Wallet', 'Payer', 'Type', 'Our_Number', 'Your_Number', 'Due_Date', 'Receiving_Agency', 'Initial_Value', 'Operations_Description', 'Operations_Value', 'Final_Value']

    df = pd.read_excel(path, skiprows=36, usecols='A:K')

    df.columns = columns

    df_filtered = df.loc[df['Wallet'] == 157][['Wallet', 'Payer', 'Type', 'Our_Number', 'Your_Number', 'Due_Date', 'Initial_Value', 'Operations_Description', 'Final_Value']]

    names = []

    file_name = os.path.basename(path)

    wallet = file_name.split("_")[2]

    date = file_name.split("_")[-1]  # returns 18-03-2024.xlsx
    date = date.split(".")[0]  # removes xlsx

    for i in df_filtered.itertuples():

        your_number = i.Your_Number

        due_date = i.Due_Date

        description = i.Operations_Description

        if description == 'settlement': 
            sql_each_your_number = f'''
                        SELECT  
                            Payer_Name as 'payer_name',
                            REPLACE(CAST(Original_Value AS VARCHAR), ',', '.') as 'original_value',
                            Document_Number as 'document_number',
                            IIF(LEFT(Payer_ID, 3) = '000', RIGHT(Payer_ID, 11), Payer_ID) as 'payer_id',
                            CAST(Payment_Date AS date) as 'payment_date',
                            CAST(Date AS date) as 'date',
                            CAST(Title_Date AS date) as 'title_date',
                            CAST(Due_Date AS date) AS 'due_date',
                            costumer_Name as 'costumer_name',
                            Type as 'type'
                        from [DW]..[vw_payer] 
                        where Document_Number = '{your_number}'
                        and Type in ('PUBLIC', 'PRIVATE')
                    '''  # column "your number" from the file should match Document_Number from the database

            result = query_sql_df(sql=sql_each_your_number, db='DW', connx = DW_SQL_CONN)

            if result.empty:  # if there are no records of titles with "your number" and it's a settlement
                names.append((i.Payer, due_date, i.Final_Value))

    names = list(set(names))
    path_generated = ''  # declare empty variable for future verification if it remains empty

    if names:  # If there are discrepancies, send by email

        file = pd.DataFrame(names, columns=['Payer Name', 'Due Date', 'Initial Value'])

        file.to_excel(f'Bank_{date}_Wallet_{wallet}.xlsx')

        path_generated = os.path.abspath(f'Bank_{date}_Wallet_{wallet}.xlsx')

        email_recipients = ['any_email_to_send@company.com']

        for email in email_recipients:
            send_email(p_from='osirisdatantech@gmail.com', p_to=email, p_subject=f'Receivable Bills Bank Wallet {wallet}', p_text=f'FOLLOWING DISCREPANCIES OF THE INDIVIDUAL BILLS OF BANK RELATED TO WALLET {wallet}',
                          p_html='<p></p>', p_password='password', p_p_filename=path_generated)


    else:
        pass

    script_directory = os.path.dirname(os.path.realpath(__file__))

    backup_folder = os.path.join(script_directory, "BACKUP")  # name of the subfolder to where the files will be moved

    if not os.path.exists(backup_folder):  # create the folder if it has been deleted
        os.makedirs(backup_folder)

    sleep(2)

    # os.remove(path) # delete the downloaded file from Bank
    os.remove(path_generated)  # delete the generated file for email sending

    sleep(5)


def send_email(email_from, email_to, subject, text, html,p_filename = None):

    try:

        msg = MIMEMultipart('multipart')
        msg['Subject'] = subject
        msg['From'] = email_from
        msg['To'] = email_to

        text = text
        html = html

        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')

        msg.attach(part1)
        msg.attach(part2)

        if p_filename is not None:
            
            if not os.path.isfile(p_filename):
                return
            ctype, encoding = mimetypes.guess_type(p_filename)

            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'

            maintype, subtype = ctype.split('/', 1)

            if maintype == 'text':
                with open(p_filename) as f:
                    mime = MIMEText(f.read(), _subtype=subtype)
            elif maintype == 'image':
                with open(p_filename, 'rb') as f:
                    mime = MIMEImage(f.read(), _subtype=subtype)
            elif maintype == 'audio':
                with open(p_filename, 'rb') as f:
                    mime = MIMEAudio(f.read(), _subtype=subtype)
            else:
                with open(p_filename, 'rb') as f:
                    mime = MIMEBase(maintype, subtype)
                    mime.set_payload(f.read())

                encoders.encode_base64(mime)

            mime.add_header('Content-Disposition', 'attachment', p_filename=p_filename.split("\\")[-1])
            #mime.add_header('Content-Disposition',f'attachment; p_filename={p_filename}',)
            msg.attach(mime)

            context = ssl.create_default_context()

            with smtplib.SMTP("email-smtp.sa-east-1.amazonaws.com", 25) as server:
                server.starttls(context=context)
                server.ehlo()
                server.login('email_login', 'email_password')
                response = server.sendmail(email_from, email_to, msg.as_string())
                server.quit()

                if not response:
                    return {"msg": "E-mail sent succesfully"}
                else:
                    return response
            
    except Exception as e:
        return e


def query_sql_df(sql,db,connx):
    try:
        conn = pymssql.connxect(server=connx['DBIP'], user=connx['DBLOGIN'], password=connx['DBPASSWORD'] , database=db)
        df = pd.read_sql_query(sql, connx)
        return df
        conn.commit()
    except Exception as e: 
        conn.rollback()
        return  None
    finally:
        conn.close()  


def main():
        
        run_download_bills_RPA(LOGIN = LOGIN, PASSWORD =  PASSWORD)
        sleep(4)
        verify_autenticy_in_database(file_path = get_file_path())


# main()