from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time


reservation_date_from = "2020/06/05"
reservation_date_to = "2020/06/05"


def send_reservation_date(from_, to_):
    driver.implicitly_wait(10)
    driver.find_element_by_id("reservation_date_from").clear()
    driver.find_element_by_id("reservation_date_from").send_keys(from_)
    driver.find_element_by_id("reservation_date_to").clear()
    driver.find_element_by_id("reservation_date_to").send_keys(to_)


def get_reserver_name(i):
    reserver_name = driver.find_element_by_xpath(f'//*[@id="row{i}resv_target_list"]/div[7]/div').text
    return reserver_name


def get_reserver_account_number(i):
    reserver_account_number = driver.find_element_by_xpath(f'//*[@id="row{i}resv_target_list"]/div[6]/div').text
    return reserver_account_number


def get_book_title(i):
    book_title = driver.find_element_by_xpath(f'//*[@id="row{i}resv_target_list"]/div[4]/div').text
    return book_title


def get_regi_number():
    book_regi_num = driver.find_element_by_id('view_reg_no').text
    return book_regi_num


def get_reserver():
    i = 0
    while True:
        try:
            reserver_name_list.append(get_reserver_name(i))
            reserver_account_list.append(get_reserver_account_number(i))
            book_title_list.append(get_book_title(i))
        except NoSuchElementException:
            break
        i += 1

    return reserver_name_list, reserver_account_list, book_title_list


def get_book_id():
    driver.current_url
    result = driver.page_source
    bs_obj = BeautifulSoup(result, "html.parser")
    table = bs_obj.find("table", id="table_contents")
    checkboxes = table.find_all('input', class_="cbox checkbox")

    return checkboxes[0].attrs['id']


def cancel_and_check_out():

    i = 0
    while True:
        try:
            number_input = driver.find_element_by_id("main_number_txt")
            number_input.clear()
            number_input.send_keys(reserver_account_list[i])
            number_input.send_keys(Keys.ENTER)
            time.sleep(1)

            reserved_booklist = driver.find_element_by_xpath('//*[@id="local_loan_container"]/ul/li[2]/a')
            click(reserved_booklist)

            check_reserved_books = driver.find_element_by_xpath('/html/body/div[1]/div/div[3]/main/div/ul[2]/li[1]/div/div[2]/div/div[1]/div[3]/div[3]/div/table/tbody/tr[2]/td[2]/input')
            click(check_reserved_books)

            detail_status = driver.find_element_by_id('bookingStatusButton')
            click(detail_status)
            time.sleep(1)

            book_regi_number_list.append(get_regi_number())
            driver.implicitly_wait(5)

            cancel_reservation = driver.find_element_by_xpath('//*[@id="common_popup_1"]/div/div/div[3]/div/div[2]/button[1]')
            click(cancel_reservation)
            time.sleep(1)

            cancel = driver.find_element_by_xpath('//*[@id="msg_btn_1"]')
            click(cancel)
            time.sleep(1)

            confirm = driver.find_element_by_xpath('//*[@id="msg_btn_1"]')
            click(confirm)
            print(f'"{book_title_list[i]}" 예약 취소')

            close = driver.find_element_by_xpath('//*[@id="common_popup_1"]/div/div/div[3]/div/div[2]/button[2]')
            click(close)
            time.sleep(1)

            regi_input = driver.find_element_by_id("main_number_txt")
            regi_input.clear()
            regi_input.send_keys(book_regi_number_list[i])
            regi_input.send_keys(Keys.ENTER)
            print(f'"{book_title_list[i]}" 대출')

            time.sleep(1)

            loan_list = driver.find_element_by_xpath('//*[@id="loanList"]')
            click(loan_list)

            book_id.append(get_book_id())

            check = driver.find_element_by_id(book_id[i])
            click(check)
            driver.implicitly_wait(5)

            renew = driver.find_element_by_xpath('//*[@id="returnDelay"]')
            click(renew)
            print(f'"{book_title_list[i]}" 대출 연장')
            time.sleep(1)

            close_button = driver.find_element_by_xpath('//*[@id="common_popup_1"]/div/div/div[3]/button')
            click(close_button)

            time.sleep(3)

        except IndexError:
            break
        i += 1


def print_receipt():

    i = 0
    while True:
        try:
            number_input = driver.find_element_by_id("main_number_txt")
            number_input.clear()
            number_input.send_keys(unique_reserver_account_list[i])
            number_input.send_keys(Keys.ENTER)
            driver.implicitly_wait(5)


            print_receipt = driver.find_element_by_id("receiptPrintBtn_main")
            click(print_receipt)
            print(f'{reserver_name_list[i]}님의 현황 확인증 출력')


        except IndexError:
            break
        i += 1


def click(x):
    driver.execute_script("arguments[0].click();", x)
    driver.implicitly_wait(5)


def get_reservation_data(i):

    dictionary = {}
    dictionary["이름"] = reserver_name_list[i]
    dictionary["회원 번호"] = reserver_account_list[i]
    dictionary["도서 제목"] = book_title_list[i]
    dictionary["도서 등록 번호"] = book_regi_number_list[i]
    dictionary["도서 id"] = book_id[i]

    return dictionary

"""------------------Alpas Login---------------"""

# Login
driver = webdriver.Chrome("C:\\Users\\user\\PycharmProjects\\untitled\\chromedriver.exe")
driver.implicitly_wait(3)
driver.get("http://152.99.43.46:28180/METIS/")
driver.maximize_window()
driver.find_element_by_name("main_login_user_id").send_keys("안병덕")
driver.find_element_by_id("user_pw").send_keys("dksqudejr1120!")
driver.find_element_by_xpath("/html/body/div[2]/div/button").click()
time.sleep(3)

# Close Pop up
popup = driver.find_element_by_xpath('//*[@id="closeNoticePopup"]')
click(popup)
driver.implicitly_wait(5)

# Open Reservation tap
navigation = driver.find_element_by_xpath('//*[@id="right_container_wrapper"]/header/div[5]/nav/a')
click(navigation)
manage_reservation = driver.find_element_by_xpath('//*[@id="COLI_01_HTML"]/li[6]')
click(manage_reservation)
driver.switch_to.window(driver.window_handles[1])
driver.implicitly_wait(5)

# Get Account Number
send_reservation_date(reservation_date_from, reservation_date_to)  # Enter reservation date
search = driver.find_element_by_xpath('//*[@id="btn_search"]')  # Search
click(search)
time.sleep(1)

# Save Reserver Data
reserver_name_list = []
reserver_account_list = []
unique_reserver_account_list = []
book_title_list = []
book_regi_number_list = []
book_id = []

# Get Reserver Data
get_reserver()

# Open Check Out tap
driver.find_element_by_xpath('//*[@id="right_container_wrapper"]/header/div[5]/nav/a').click()
driver.find_element_by_xpath('//*[@id="COLI_01_HTML"]/li[2]').click()
driver.switch_to.window(driver.window_handles[2])
driver.implicitly_wait(3)

# Show Only Local Library Data
only_local = driver.find_element_by_id("is_only_local")
click(only_local)
driver.implicitly_wait(3)


# Set Unique Account List
for v in reserver_account_list:
    if v not in unique_reserver_account_list:
        unique_reserver_account_list.append(v)


# Get Registration Number, Cancel, and Check Out
cancel_and_check_out()
print_receipt()
