import xlrd
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

class Read_Ex():
    def read_excel(self):
        #打开excel表，填写路径
        book = xlrd.open_workbook("/Users/qy/Downloads/学生工作/就业方案2.xlsx")
        #找到sheet页
        table = book.sheet_by_name("2")
        #获取总行数总列数
        row_Num = table.nrows
        col_Num = table.ncols

        s =[]
        key =table.row_values(0)# 这是第一行数据，作为字典的key值

        if row_Num <= 1:
            print("没数据")
        else:
            j = 1
            for i in range(row_Num-1):
                d ={}
                values = table.row_values(j)
                for x in range(col_Num):
                    # 把key值对应的value赋值给key，每行循环
                    d[key[x]]=values[x]
                j+=1
                # 把字典加到列表中
                s.append(d)
            return s

def openChrome():
    # 加启动配置
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')#有界面模式
    #option.add_argument('--headless')#无界面模式
    # 打开chrome浏览器
    driver = webdriver.Chrome(options=option)
    return driver

def operationAuth(driver,list):
    
    url = "https://jy.ncss.org.cn"   #地址
    driver.get(url)
    
    # 找到输入框并输入查询内容
    elem=WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_id("username"))
    elem.send_keys('123355jy0211')   #账号
    elem = WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_id("password"))
    elem.send_keys('2231611jy')  #密码
   
    # 提交表单
    WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("//*[@id='logintype']")).click()
    WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("//*[@name='submit']")).click()
    
    WebDriverWait(driver,15,0.5).until(lambda x:x.find_element_by_link_text("学生列表")).click()
    
    for i in list:
        driver.get("https://jy.ncss.org.cn/graduate/2020/list.html")
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("//*[@id='keyword']")).clear()
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_class_name("form-control")).send_keys(i['name'])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("//*[@id='search-btn']")).click()
        time.sleep(1)
        WebDriverWait(driver,15,0.5).until(lambda x:x.find_element_by_link_text("毕业去向")).click()
        driver.get("https://jy.ncss.org.cn/graduate/2020/list.html#sdc-byqx")
        WebDriverWait(driver,15,0.5).until(lambda x:x.find_element_by_link_text("点击修改")).click()
    
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_class_name("select2-chosen")).click()
        time.sleep(2)
        jylx=i['jylx']
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(jylx[0:2])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(Keys.ENTER)
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(i['dwname'])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[2]/td[1]/span[2]/input")).send_keys(i['code'])
    
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_id("s2id_dwxz")).click()#单位性质
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(i['dwxz'])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(Keys.ENTER)

        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_id("s2id_dwhy")).click()#单位行业
        time.sleep(1)
        hylb=i['hylb']
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(hylb[0:1])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(Keys.ENTER)

        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_id("s2id_dwszd")).click()#单位所在地
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(i['dwdz'])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(Keys.ENTER)
   
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_id("s2id_gzzwlb")).click()#工作类别
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys("其他")
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[1]/td[2]/span[2]/input")).send_keys(Keys.ENTER)
        time.sleep(1)

        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[4]/td[2]/span[2]/input")).send_keys(i['dwlxr'])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[5]/td[1]/span[2]/input")).send_keys(i['dwphone'])
        time.sleep(1)
        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/table[1]/tbody/tr[7]/td[1]/span[2]/input")).send_keys(i['dwdz'])
        time.sleep(1)

        WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/button[1]")).click()#点击确认按钮
        time.sleep(2)

        dig_alert = driver.switch_to.alert
        time.sleep(1)
# 打印警告对话框内容
        print(i['name']+dig_alert.text)
# alert对话框属于警告对话框，我们这里只能接受弹窗
        dig_alert.accept()

    #WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_css_selector('#s2id_autogen50_search')).send_keys("其他录用形式就业")

   
    """
    document.querySelector("#s2id_autogen50_search")
    //*[@id="s2id_autogen50_search"]
    WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("select2-input")).send_keys("其他录用形式就业")
    WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("select2-input")).send_keys(Keys.ENTER)
    time.sleep(5)
    """
    #WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_class_name("employmentInfo.dwmc")).send_keys("申晓倩")
    
    """
    WebDriverWait(driver,30,0.5).until(lambda x:x.find_element_by_xpath("//*[@id='f7']")).click()
    driver.implicitly_wait(30)
    driver.switch_to.frame("tabs_7_iframe")
    #driver.find_element_by_link_text("上班签到").click()
    try:
         WebDriverWait(driver,15,0.5).until(lambda x:x.find_element_by_link_text("上班签到")).click()
         success=success+"; "+idNo
    except:
        print(idNo)
        failed=failed+"; "+idNo
    """
if __name__ == '__main__':
    
    driver = openChrome()
    r = Read_Ex()
    s=r.read_excel()
    operationAuth(driver,s)
    driver.quit() 

    """
    r = Read_Ex()
    s=r.read_excel()
    for i in s:
        print(i)
        print('\n')
   """