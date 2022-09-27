from selenium import webdriver
import time
import pandas as pd
import xlsxwriter
import pathlib

def GetFileExtension(file_name):
    file_extension = pathlib.Path(file_name).suffix
    return file_extension

driver = webdriver.Chrome(executable_path="Duong dan den file webdriver")

driver.set_window_size(1000, 1000)

driver.get('link bai tap')
driver.find_element_by_id('username').send_keys('Tai khoan cst')
driver.find_element_by_id('password').send_keys('Mat khau cst')
driver.find_element_by_id('loginbtn').click()
time.sleep(5)

list_a = []

slsv = 50
i = 1
while(i<=slsv):
    time.sleep(3)
    trs = driver.find_elements_by_xpath('/html/body/div[3]/div[5]/div/div/section/div[1]/div[4]/div[3]/table/tbody/tr[%i]' %i)
    for j in trs:
        time.sleep(3)
        count = 0
        list_b = []
        tds = j.find_elements_by_tag_name('td')
        for td in tds:
            count+=1
            if(count==3 or count==4):
                value = td.text
                list_b.append(value)
                print('Thong tin: ', value)
            elif(count==10):
                if('tệp' in td.text):
                    td.find_element_by_tag_name('a').click()
                    count_file = 0
                    for file in driver.find_elements_by_class_name('fileuploadsubmission'):
                        get_file_name = file.find_element_by_tag_name('a').text
                        file_extension = GetFileExtension(get_file_name)
                        if(file_extension == '.html'or file_extension == '.css' or file_extension == '.js' or file_extension == '.docx' or file_extension == '.doc'):
                            count_file+=1
                    list_b.append(count_file)
                    print('File', count_file)
                    driver.back()
                else:
                    count_file = 0
                    for file in td.find_elements_by_class_name('fileuploadsubmission'):
                        get_file_name = file.find_element_by_tag_name('a').text
                        file_extension = GetFileExtension(get_file_name)
                        if(file_extension == '.html'or file_extension == '.css' or file_extension == '.js' or file_extension == 'docx' or file_extension == 'doc'):
                            count_file+=1
                    list_b.append(count_file)
                    print('File', count_file)
    print('i = ', i)
    i+=1
    list_a.append(list_b)
header = ["Họ và tên", "Mã sv", "Số file"]
df = pd.DataFrame(list_a, columns=header)
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='welcome', index=False)
writer.save()
time.sleep(3)
driver.quit()
