import pandas as pd
import time
import pprint as pp
from openpyxl import load_workbook
from selenium import webdriver as wb
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl.worksheet.properties import WorksheetProperties as wp
from selenium.webdriver.common.action_chains import ActionChains

# List of product skus
sku_list = [
'List of product skus here'
]

# Url to use
url_path = 'url_here'


# Setting up the webdriver for Selenium
options = wb.ChromeOptions()
options.add_argument('--start-maximized')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = wb.Chrome(options=options)

#  Creating driver
driver.get(url_path)
time.sleep(1)
    
# Setting up a dictionarys to use for the data
src_dict = {'Sku': [], 'Img_url1': [], 'Img_url2': [], 'Img_url3': [], 'Img_url4': [], 'Specs_Url': [], 'Installation_Url': [], 'Skus_Not_Found': []}
link_dict = {'Install': ['-installation-sheet.pdf', '_installation_sheet.pdf', 'install.pdf'], 'Specs': ['specification-sheet.pdf', 'specification_sheet.pdf', 'spec.pdf']}

# Checking the first sku in the list
for sku in sku_list:
    if sku == sku_list[0]:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'term')))
        driver.find_element(By.NAME, "term").send_keys(sku)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/header/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div[1]/a/div[1]/img')))
        driver.find_element(By.NAME, "term").send_keys(Keys.RETURN)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class,'aspect-ratio--object z-1')]")))
        

        try:
            # If the text element doesn't match the sku from the list then append to dictionary and move on
            if driver.find_element_by_xpath("//*[contains(@class, 'f6 mt1 lh-title theme-grey-medium truncate')]").text != 'Model: ' + sku:
                
                if driver.find_element_by_xpath("//div[contains(@class, 'ku8y0w3')]"):
                    driver.find_element_by_xpath("//div[contains(@class, 'ku8y0w6')]").click()
                
                src_dict['Img_url1'].append('NULL')
                src_dict['Img_url2'].append('NULL')
                src_dict['Img_url3'].append('NULL')
                src_dict['Img_url4'].append('NULL')
                src_dict['Sku'].append(sku)
                src_dict['Skus_Not_Found'].append(sku)

            # If the text matches the sku move forward 
            elif driver.find_element_by_xpath("//*[contains(@class, 'f6 mt1 lh-title theme-grey-medium truncate')]").text == 'Model: ' + sku:
                driver.find_element_by_xpath("//*[contains(@class,'aspect-ratio--object z-1')]").click()
                time.sleep(3)
                
                # If the banner element is found close it
                if driver.find_element_by_xpath("//div[contains(@class, 'ku8y0w3')]"):
                    driver.find_element_by_xpath("//div[contains(@class, 'ku8y0w6')]").click()
                else:
                    continue
                
                # Extracting the image src
                src_1 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                time.sleep(1)
                src_dict['Img_url1'].append(src_1)
                
                # Trying to extract the next 3 img src or put null in the dictionary
                try:
                    driver.find_element_by_xpath("//*[contains(@aria-label,'thumb slide 1')]").click()
                    time.sleep(1)
                    src_2 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                    src_dict['Img_url2'].append(src_2)
                
                except:
                    src_dict['Img_url2'].append('NULL')
                
                try:
                    driver.find_element_by_xpath("//*[contains(@aria-label,'thumb slide 2')]").click()
                    time.sleep(1)
                    src_3 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                    src_dict['Img_url3'].append(src_3)
                
                except:
                    src_dict['Img_url3'].append('NULL')

                try:  
                    driver.find_element_by_xpath("//*[contains(@aria-label,'thumb slide 3')]").click()
                    time.sleep(1)
                    src_4 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                    src_dict['Img_url4'].append(src_4)
                    src_dict['Sku'].append(sku)
                    pp.pprint(src_dict)
                
                except:
                    src_dict['Img_url4'].append('NULL')
                    src_dict['Sku'].append(sku)
                    pp.pprint(src_dict)

                # Extracting the pdf hrefs, sorting them, and appending them to the dictionary
                elements = driver.find_elements_by_xpath("//a[contains(@class, 'f-inherit fw-inherit link theme-primary  pb3 f7 db underline-hover')]")
                pdf_links = [element.get_attribute('href') for element in elements]
                
                for link in pdf_links:
                    if link.endswith(tuple(link_dict['Install'])):
                        src_dict['Installation_Url'].append(link)
                    elif link.endswith(tuple(link_dict['Specs'])):
                        src_dict['Specs_Url'].append(link)
                    else:
                        src_dict['Specs_Url'].append('NULL')
                        src_dict['Installation_Url'].append('NULL')


        except:
            driver.find_element(By.NAME, "term").clear()
            src_dict['Img_url1'].append('NULL')
            src_dict['Img_url2'].append('NULL')
            src_dict['Img_url3'].append('NULL')
            src_dict['Img_url4'].append('NULL')
            src_dict['Sku'].append(sku)
            src_dict['Skus_Not_Found'].append(sku)
            pp.pprint(src_dict)
            
    if sku != sku_list[0]:
        driver.find_element(By.NAME, "term").clear()
        break

# Checking all skus after the first sku
for sku in sku_list[1:]:
    driver.find_element(By.NAME, "term").click()
    driver.find_element(By.NAME, "term").send_keys(sku)
    time.sleep(1)
    driver.find_element(By.NAME, "term").send_keys(Keys.RETURN)
    time.sleep(2)
    
    try:
        # If the text element doesn't match the sku from the list then append to dictionary and move on
        if driver.find_element_by_xpath("//*[contains(@class, 'f6 mt1 lh-title theme-grey-medium truncate')]").text != 'Model: ' + sku:
            driver.find_element(By.NAME, "term").clear()
            src_dict['Img_url1'].append('NULL')
            src_dict['Img_url2'].append('NULL')
            src_dict['Img_url3'].append('NULL')
            src_dict['Img_url4'].append('NULL')
            src_dict['Sku'].append(sku)
            src_dict['Skus_Not_Found'].append(sku)

        # If the text matches the sku move forward 
        elif driver.find_element_by_xpath("//*[contains(@class, 'f6 mt1 lh-title theme-grey-medium truncate')]").text == 'Model: ' + sku:
            driver.find_element_by_xpath("//*[contains(@class,'aspect-ratio--object z-1')]").click()
            time.sleep(3)
            
            #  Extracting the first image src
            src_5 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
            time.sleep(1)
            src_dict['Img_url1'].append(src_5)
            
            # Trying to extract the next 3 img srcs or put null in the dictionary
            try:
                driver.find_element_by_xpath("//*[contains(@aria-label,'thumb slide 1')]").click()
                time.sleep(1)
                src_6 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                src_dict['Img_url2'].append(src_6)

            except:
                src_dict['Img_url2'].append('NULL')
                
            try:    
                driver.find_element_by_xpath("//*[contains(@aria-label,'thumb slide 2')]").click()
                time.sleep(1)
                src_7 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                src_dict['Img_url3'].append(src_7)

            except:
                src_dict['Img_url3'].append('NULL')

            try:   
                driver.find_element_by_xpath("//*[contains(@aria-label,'thumb slide 3')]").click()
                time.sleep(1)
                src_8 = driver.find_element_by_xpath("//*[contains(@class,'w-auto self-center undefined')]").get_attribute('src')
                src_dict['Img_url4'].append(src_8)
                src_dict['Sku'].append(sku)
                print(src_dict)

            except:
                src_dict['Img_url4'].append('NULL')
                src_dict['Sku'].append(sku)
                pp.pprint(src_dict)
                
            if str(src_2).endswith('noimage.gif'):
                src_dict['Img_url'].append('NULL')
            else:
                src_2.get_attribute('src')
            
            # Extracting the pdf hrefs, sorting them, and appending them to the dictionary
            elements = driver.find_elements_by_xpath("//a[contains(@class, 'f-inherit fw-inherit link theme-primary  pb3 f7 db underline-hover')]")
            pdf_links = [element.get_attribute('href') for element in elements]
            
            for link in pdf_links:
                if link[-23:] ==  'specification-sheet.pdf' or link[-23:] ==  'specification_sheet.pdf' or link[-8:] == 'spec.pdf':
                    src_dict['Specs_Url'].append(link)
                elif link[-23:] == '-installation-sheet.pdf' or link[-23:] == '_installation_sheet.pdf' or link[-11:] == 'install.pdf':
                    src_dict['Installation_Url'].append(link)
                if link.endswith(tuple(link_dict['Install'])):
                    src_dict['Installation_Url'].append(link)
                elif link.endswith(tuple(link_dict['Specs'])):
                    src_dict['Specs_Url'].append(link)
                else:
                    src_dict['Specs_Url'].append('NULL')
                    src_dict['Installation_Url'].append('NULL')

    except:
        driver.find_element(By.NAME, "term").clear()
        src_dict['Img_url1'].append('NULL')
        src_dict['Img_url2'].append('NULL')
        src_dict['Img_url3'].append('NULL')
        src_dict['Img_url4'].append('NULL')
        src_dict['Sku'].append(sku)
        src_dict['Skus_Not_Found'].append(sku)
        pp.pprint(src_dict)

# Quiting the driver and manipulation the dictionary into a dataframe 
driver.quit()
df = pd.DataFrame.from_dict(src_dict,orient='index')
df = df.transpose()

# Writing the dataframe to an excel worksheet
path = 'excel_file.xlsx'
excel_wb = load_workbook(path)
with pd.ExcelWriter(path) as writer:
    writer.book = excel_wb
    df.to_excel(writer, sheet_name='Asset_Data', index=False)
    file_sheet = writer.sheets['Asset_Data']
    file_sheet.sheet_properties.tabColor = 'FFFF00'
    
print('Run Complete!')
