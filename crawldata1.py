from selenium import webdriver
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import openpyxl

def get_value_excel (filename, cellname):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet']
    wb.close()
    return Sheet1[cellname].value
def update_value_excel(filename, cellname, value):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet']
    Sheet1[cellname].value = value
    wb.close()
    wb.save(filename)    

col_name_link = "A"
col_name_id = "B"
col_name_title = "C"
col_name_time= "D"
col_name_ngayhethan= "E"
col_name_diachi= "F"
col_name_gia= "G"
col_name_dientich= "H"
col_name_agency= "I"
col_name_profilelink= "J"
filename = 'C://Users//ASUS//crawldata//crawldata//spiders//files.xlsx'
    
for i_row in range(2,182):
    cell_name_link = "%s%s"%(col_name_link,i_row)
    cell_name_id = "%s%s"%(col_name_id,i_row)
    cell_name_title = "%s%s"%(col_name_title,i_row)
    cell_name_time = "%s%s"%(col_name_time,i_row)
    cell_name_ngayhethan = "%s%s"%(col_name_ngayhethan,i_row)
    cell_name_diachi = "%s%s"%(col_name_diachi,i_row)
    cell_name_gia = "%s%s"%(col_name_gia,i_row)
    cell_name_dientich = "%s%s"%(col_name_dientich,i_row)
    cell_name_agency = "%s%s"%(col_name_agency,i_row)
    cell_name_profilelink = "%s%s"%(col_name_profilelink,i_row)
    link = get_value_excel (filename, cell_name_link)
    driver = webdriver.Chrome("C:\\Users\\ASUS\\crawldata\\crawldata\\spiders\\chromedriver.exe")
    driver.get("https://homedy.com" + link)
    sleep(5)
    # mstSearch = driver.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input")
    # sleep(1)
    # mstSearch.send_keys(link)
    # mstSearchBtn = driver.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]")
    # sleep(1)
    # mstSearchBtn.click()
    # sleep(1)  
    mstSearchid = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/div[2]/span[2]")
    sleep(1)
    mstSearchtitle = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/h1")
    mstSearchtime = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/div[2]/span[5]")
    mstSearchngayhethan = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/div[2]/span[7]")
    mstSearchdiachi = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/div[3]")
    mstSearchgia = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/div[4]/div[1]/div[1]/strong")
    mstSearchdientich = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[1]/div/div[1]/div/div[4]/div[1]/div[2]/strong")
    mstSearchagency = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div/a[1]")
    mstSearchprofilelink = driver.find_element_by_xpath("/html/body/div[1]/div[5]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/div/a[1]").get_attribute('href')
    sleep(1)
    update_value_excel(filename,cell_name_id,mstSearchid.text)
    update_value_excel(filename,cell_name_title,mstSearchtitle.text)
    update_value_excel(filename,cell_name_time,mstSearchtime.text)
    update_value_excel(filename,cell_name_ngayhethan,mstSearchngayhethan.text)
    update_value_excel(filename,cell_name_diachi,mstSearchdiachi.text)
    update_value_excel(filename,cell_name_gia,mstSearchgia.text)
    update_value_excel(filename,cell_name_dientich,mstSearchdientich.text)
    update_value_excel(filename,cell_name_agency,mstSearchagency.text)
    update_value_excel(filename,cell_name_profilelink,mstSearchprofilelink)
    driver.close()
    sleep(1)
        

