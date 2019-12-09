from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import time
import random

monthly_stat_from = "2019/12/01"
monthly_stat_to = "2019/12/31"

loan_date_from = "2019/12/02"
loan_date_to = "2019/12/06"

wb = load_workbook("C:\\Users\\user\\Desktop\\주간보고\\주간보고(2019.12.02~2019.12.06).xlsx")
ws = wb["주간업무일반"]

def send_checkout_date(from_, to_):
    driver.implicitly_wait(10)
    driver.find_element_by_id("loan_date_from").clear()
    driver.find_element_by_id("loan_date_from").send_keys(from_)
    driver.find_element_by_id("loan_date_to").clear()
    driver.find_element_by_id("loan_date_to").send_keys(to_)

def send_return_date(from_, to_):
    driver.implicitly_wait(10)
    driver.find_element_by_id("return_date_from").clear()
    driver.find_element_by_id("return_date_from").send_keys(from_)
    driver.find_element_by_id("return_date_to").clear()
    driver.find_element_by_id("return_date_to").send_keys(to_)

def search_and_close():
    driver.implicitly_wait(10)
    search = driver.find_element_by_xpath('//*[@id="statistics_search_btn"]')
    driver.execute_script("arguments[0].click();", search)
    driver.implicitly_wait(10)
    driver.find_element_by_xpath('//*[@id="msg_btn_1"]').click()
    driver.implicitly_wait(10)
    close_button = driver.find_element_by_xpath('//*[@id="pusrchaesDataPrintdeSearchClose"]')
    driver.execute_script("arguments[0].click();", close_button)

def get_page_info():
    result = driver.page_source
    bs_obj = BeautifulSoup(result, "html.parser")
    table = bs_obj.find("table", class_="display compact cell-border table table-bordered ui-jqgrid-btable ui-common-table")
    tr = table.find_all("tr")

    general_works = tr[1].find_all("td")
    philosophy = tr[2].find_all("td")
    religion = tr[3].find_all("td")
    social_sciences = tr[4].find_all("td")
    natural_sciences = tr[5].find_all("td")
    technology = tr[6].find_all("td")
    arts = tr[7].find_all("td")
    language = tr[8].find_all("td")
    literature = tr[9].find_all("td")
    history = tr[10].find_all("td")

    dictionary = {}
    dictionary["general_works"] = [int(general_works[i].text) for i in range(4, 9)]
    dictionary["philosophy"] = [int(philosophy[i].text) for i in range(4, 9)]
    dictionary["religion"] = [int(religion[i].text) for i in range(4, 9)]
    dictionary["social_sciences"] = [int(social_sciences[i].text) for i in range(4, 9)]
    dictionary["natural_sciences"] = [int(natural_sciences[i].text) for i in range(4, 9)]
    dictionary["technology"] = [int(technology[i].text) for i in range(4, 9)]
    dictionary["arts"] = [int(arts[i].text) for i in range(4, 9)]
    dictionary["language"] = [int(language[i].text) for i in range(4, 9)]
    dictionary["literature"] = [int(literature[i].text) for i in range(4, 9)]
    dictionary["history"] = [int(history[i].text) for i in range(4, 9)]
    
    dictionary["monthly_general_works"] = int(general_works[3].text)
    dictionary["monthly_philosophy"] = int(philosophy[3].text)
    dictionary["monthly_religion"] = int(religion[3].text)
    dictionary["monthly_social_sciences"] = int(social_sciences[3].text)
    dictionary["monthly_natural_sciences"] = int(natural_sciences[3].text)
    dictionary["monthly_technology"] = int(technology[3].text)
    dictionary["monthly_arts"] = int(arts[3].text)
    dictionary["monthly_language"] = int(language[3].text)
    dictionary["monthly_literature"] = int(literature[3].text)
    dictionary["monthly_history"] = int(history[3].text)

    return dictionary

"""------------------Alpas Login---------------"""

#Login
driver = webdriver.Chrome("C:\\Users\\user\\PycharmProjects\\untitled\\chromedriver.exe")
driver.implicitly_wait(3)
driver.get("http://152.99.43.46:28180/METIS/")
driver.maximize_window()
driver.find_element_by_name("main_login_user_id").send_keys("안병덕")
driver.find_element_by_id("user_pw").send_keys("dksqudejr1120!")
driver.find_element_by_xpath("/html/body/div[2]/div/button").click()

#Open Statistics tap
driver.find_element_by_xpath('//*[@id="right_container_wrapper"]/header/div[5]/nav/a').click()
driver.find_element_by_xpath('//*[@id="COLI_01_HTML"]/li[12]').click()
driver.switch_to.window(driver.window_handles[1])

"""------------------Checkout Statistics---------------"""

driver.find_element_by_xpath('//*[@id="statistics_type"]/option[6]').click()                                            #Choose Checkout Statistics
driver.find_element_by_xpath('//*[@id="statistics_btn"]').click()                                                       #Click Statistics button
send_checkout_date(loan_date_from, loan_date_to)                                                                        #Enter checkout date
driver.find_element_by_xpath('//*[@id="shelf_loc_code_list"]/option[34]').click()                                       #Choose Songdo Library of International Organization
driver.find_element_by_xpath('//*[@id="shelf_loc_code_btn"]').click()
driver.implicitly_wait(10)
by_date = driver.find_element_by_xpath('//*[@id="jqg82"]/td[2]')                                                        #Adjust matrix
driver.execute_script("arguments[0].click();", by_date)
search_and_close()                                                                                                      #Search
get_page_info()                                                                                                         #Parsing


#Save in Excel
ws["F43"], ws["F44"], ws["F45"], ws["F46"], ws["F47"] = [get_page_info()["general_works"][i] for i in range(5)]
ws["G43"], ws["G44"], ws["G45"], ws["G46"], ws["G47"] = [get_page_info()["philosophy"][i] for i in range(5)]
ws["H43"], ws["H44"], ws["H45"], ws["H46"], ws["H47"] = [get_page_info()["religion"][i] for i in range(5)]
ws["I43"], ws["I44"], ws["I45"], ws["I46"], ws["I47"] = [get_page_info()["social_sciences"][i] for i in range(5)]
ws["J43"], ws["J44"], ws["J45"], ws["J46"], ws["J47"] = [get_page_info()["natural_sciences"][i] for i in range(5)]
ws["K43"], ws["K44"], ws["K45"], ws["K46"], ws["K47"] = [get_page_info()["technology"][i] for i in range(5)]
ws["L43"], ws["L44"], ws["L45"], ws["L46"], ws["L47"] = [get_page_info()["arts"][i] for i in range(5)]
ws["M43"], ws["M44"], ws["M45"], ws["M46"], ws["M47"] = [get_page_info()["language"][i] for i in range(5)]
ws["N43"], ws["N44"], ws["N45"], ws["N46"], ws["N47"] = [get_page_info()["literature"][i] for i in range(5)]
ws["O43"], ws["O44"], ws["O45"], ws["O46"], ws["O47"] = [get_page_info()["history"][i] for i in range(5)]

time.sleep(3)

"""------------------Cumulative Checkout Statistics---------------"""

driver.find_element_by_xpath('//*[@id="statistics_btn"]').click()
send_checkout_date(monthly_stat_from, monthly_stat_to)
search_and_close()
get_page_info()

#Save in Excel
ws["F49"] = get_page_info()["monthly_general_works"]
ws["G49"] = get_page_info()["monthly_philosophy"]
ws["H49"] = get_page_info()["monthly_religion"]
ws["I49"] = get_page_info()["monthly_social_sciences"]
ws["J49"] = get_page_info()["monthly_natural_sciences"]
ws["K49"] = get_page_info()["monthly_technology"]
ws["L49"] = get_page_info()["monthly_arts"]
ws["M49"] = get_page_info()["monthly_language"]
ws["N49"] = get_page_info()["monthly_literature"]
ws["O49"] = get_page_info()["monthly_history"]

time.sleep(3)


"""------------------Return Statistics---------------"""

driver.find_element_by_xpath('//*[@id="statistics_type"]/option[8]').click()
driver.find_element_by_xpath('//*[@id="statistics_btn"]').click()
driver.implicitly_wait(10)
checkbox = driver.find_element_by_xpath('//*[@id="loan_return_statistics_search_modal"]/div/div/div[2]/div[1]/table[2]/tbody/tr[1]/th[2]/label')
driver.execute_script("arguments[0].click();", checkbox)
send_return_date(loan_date_from, loan_date_to)
driver.implicitly_wait(10)
driver.find_element_by_xpath('//*[@id="jqg162"]/td[2]').click()
search_and_close()
get_page_info()

#Save in Excel
ws["F56"], ws["F57"], ws["F58"], ws["F59"], ws["F60"] = [get_page_info()["general_works"][i] for i in range(5)]
ws["G56"], ws["G57"], ws["G58"], ws["G59"], ws["G60"] = [get_page_info()["philosophy"][i] for i in range(5)]
ws["H56"], ws["H57"], ws["H58"], ws["H59"], ws["H60"] = [get_page_info()["religion"][i] for i in range(5)]
ws["I56"], ws["I57"], ws["I58"], ws["I59"], ws["I60"] = [get_page_info()["social_sciences"][i] for i in range(5)]
ws["J56"], ws["J57"], ws["J58"], ws["J59"], ws["J60"] = [get_page_info()["natural_sciences"][i] for i in range(5)]
ws["K56"], ws["K57"], ws["K58"], ws["K59"], ws["K60"] = [get_page_info()["technology"][i] for i in range(5)]
ws["L56"], ws["L57"], ws["L58"], ws["L59"], ws["L60"] = [get_page_info()["arts"][i] for i in range(5)]
ws["M56"], ws["M57"], ws["M58"], ws["M59"], ws["M60"] = [get_page_info()["language"][i] for i in range(5)]
ws["N56"], ws["N57"], ws["N58"], ws["N59"], ws["N60"] = [get_page_info()["literature"][i] for i in range(5)]
ws["O56"], ws["O57"], ws["O58"], ws["O59"], ws["O60"] = [get_page_info()["history"][i] for i in range(5)]

time.sleep(3)


"""------------------Cumulative Return Statistics---------------"""

driver.find_element_by_xpath('//*[@id="statistics_btn"]').click()
send_return_date(monthly_stat_from, monthly_stat_to)
search_and_close()
time.sleep(3)
get_page_info()

#Save in Excel
ws["F62"] = get_page_info()["monthly_general_works"]
ws["G62"] = get_page_info()["monthly_philosophy"]
ws["H62"] = get_page_info()["monthly_religion"]
ws["I62"] = get_page_info()["monthly_social_sciences"]
ws["J62"] = get_page_info()["monthly_natural_sciences"]
ws["K62"] = get_page_info()["monthly_technology"]
ws["L62"] = get_page_info()["monthly_arts"]
ws["M62"] = get_page_info()["monthly_language"]
ws["N62"] = get_page_info()["monthly_literature"]
ws["O62"] = get_page_info()["monthly_history"]

time.sleep(3)


"""------------------User Statistics---------------"""

driver.find_element_by_xpath('//*[@id="statistics_type"]/option[1]').click()
driver.find_element_by_xpath('//*[@id="statistics_btn"]').click()
send_checkout_date(loan_date_from, loan_date_to)
driver.implicitly_wait(10)
by_date = driver.find_element_by_xpath('//*[@id="jqg242"]/td[2]')
by_class = driver.find_element_by_xpath('//*[@id="jqg264"]/td[2]')
driver.execute_script("arguments[0].click();", by_date)
driver.execute_script("arguments[0].click();", by_class)
search_and_close()

#Parsing
result = driver.page_source
bs_obj = BeautifulSoup(result, "html.parser")
table = bs_obj.find("table", class_="display compact cell-border table table-bordered ui-jqgrid-btable ui-common-table")
tr = table.find_all("tr")

male = tr[1].find_all("td")
female = tr[2].find_all("td")
foreigner = tr[3].find_all("td")

#Save in Excel
ws["F30"], ws["F31"], ws["F32"], ws["F33"], ws["F34"] = [int(male[i].text) + int(female[i].text) for i in range(4, 9)]
ws["G30"], ws["G31"], ws["G32"], ws["G33"], ws["G34"] = [int(foreigner[i].text) for i in range(4, 9)]
ws["H30"], ws["I30"], ws["H31"], ws["I31"], ws["H32"], ws["I32"], ws["H33"], ws["I33"], ws["H34"], ws["I34"] = "=F30*2", "=G30*2", "=F31*2", "=G31*2", "=F32*2", "=G32*2", "=F33*2", "=G33*2", "=F34*2", "=G34*2"

time.sleep(3)


"""------------------Cumulative User Statistics---------------"""

driver.find_element_by_xpath('//*[@id="statistics_btn"]').click()
send_checkout_date(monthly_stat_from, monthly_stat_to)
search_and_close()

#Parsing
result = driver.page_source
bs_obj = BeautifulSoup(result, "html.parser")
table = bs_obj.find("table", class_="display compact cell-border table table-bordered ui-jqgrid-btable ui-common-table")
tr = table.find_all("tr")

male = tr[1].find_all("td")
female = tr[2].find_all("td")
"""male_youth = tr[3].find_all("td")"""
foreigner = tr[3].find_all("td")
"""disabled =  tr[5].find_all("td")"""

#Save in Excel
ws["F36"] = int(male[3].text) + int(female[3].text)
ws["G36"] = int(foreigner[3].text)

time.sleep(3)

"""------------------Save Excel---------------"""

#Get Dates
loan_date_from_month = int(loan_date_from[5:7])
loan_date_from_day = int(loan_date_from[8:10])

monday = str(loan_date_from_month) + "/" + str(loan_date_from_day)
tuesday = str(loan_date_from_month) + "/" + str(loan_date_from_day + 1)
wednesday = str(loan_date_from_month) + "/" + str(loan_date_from_day + 2)
thursday = str(loan_date_from_month) + "/" + str(loan_date_from_day + 3)
friday = str(loan_date_from_month) + "/" + str(loan_date_from_day + 4)
saturday = str(loan_date_from_month) + "/" + str(loan_date_from_day + 5)
sunday = str(loan_date_from_month) + "/" + str(loan_date_from_day + 6)

#Save in Excel
ws["A2"] = loan_date_from + "(월) ~ " + loan_date_to + "(금)"
ws["C6"], ws["A15"], ws["A30"], ws["A43"], ws["A56"] = [monday for i in range(0, 5)]
ws["E6"], ws["A16"], ws["A31"], ws["A44"], ws["A57"] = [tuesday for i in range(0, 5)]
ws["G6"], ws["A17"], ws["A32"], ws["A45"], ws["A58"] = [wednesday for i in range(0, 5)]
ws["I6"], ws["A18"], ws["A33"], ws["A46"], ws["A59"] = [thursday for i in range(0, 5)]
ws["K6"], ws["A19"], ws["A34"], ws["A47"], ws["A60"] = [friday for i in range(0, 5)]
ws["M6"] = saturday
ws["N6"] = sunday

#Random value for PC & Magazine User
ws["J30"], ws["J31"], ws["J32"], ws["J33"], ws["J34"] = [random.randint(18, 26) for i in range(0,5)]
ws["K30"], ws["K31"], ws["K32"], ws["K33"], ws["K34"] = [random.randint(0, 1) for i in range(0,5)]
ws["L30"], ws["L31"], ws["L32"], ws["L33"], ws["L34"] = [random.randint(22, 30) for i in range(0,5)]
ws["M30"], ws["M31"], ws["M32"], ws["M33"], ws["M34"] = [random.randint(0, 3) for i in range(0,5)]

wb.save("C:\\Users\\user\\Desktop\\주간보고\\test.xlsx")

