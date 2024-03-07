import zipfile

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import dataEntryFunction as df
import shutil
import time
import threading
from distutils.dir_util import copy_tree




## dataEntryFunction.py에서 set_filename()로 설정한 날짜는 오늘날짜(다운로드폴더 하위폴더 이름)입니다.
## 다운로드 대상 파일의 날짜는 직접 입력하세요. (보통 오늘날짜 하루 전)

## ex) 오늘이 2023년 6월 15일의 경우
## 1. set_filename()에서 filename = "230615" (또는 "") - 보통 수정할 일 없음
## 2. obj_dates = ["230614"] (보통 오늘날짜보다 하루 전, 월요일의 경우에는 토요일, 화요일의 경우에는 일요일, 월요일)
## 으로 입력하시면 됩니다.

today_yymmdd, _ = df.set_filename()
today_mmdd = today_yymmdd[2:6]
download_dir = "C:/Users/user/Downloads/"  # 다운로드 디렉토리 경로
obj_dates = df.getDownloadObjDates()  ## 다운로드 받을 날짜 (보통 오늘날짜보다 하루 전, 월요일의 경우에는 토요일, 화요일의 경우에는 일요일, 월요일 날짜를 넣으면 됨)

dup_check_dir_nm = 'inv_dup_check'
if not os.path.isdir(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}'):
    os.mkdir(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}')

list_past = []
list_current = []
if os.path.isfile(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}/{today_yymmdd}.txt'):
    fobj_dupchk = open(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}/{today_yymmdd}.txt', 'r')
    arr_already = fobj_dupchk.readlines()
    for item_already in arr_already:
        list_past.append(item_already.rsplit('\n', 1)[0])
    fobj_dupchk.close()
print(f'이미 받은 uri 목록:')
for i in list_past:
    print(f"\t{i}")

## 체크하는부분
print("오늘 날짜 설정")
print(f"\t{today_yymmdd} ({df.get_weekday_from_yymmdd(today_yymmdd)})")
if obj_dates == [] or obj_dates is None :
    filename_yymmdd, _ = df.set_filename()
    print("obj_dates가 비어있습니다.", (", ".join(df.getDownloadObjDates())),"날짜로 다운받으려면 아무키나 눌러주세요.")
    check = input("계속 진행하려면 아무키나 입력하세요.")
    obj_dates = df.getDownloadObjDates()
else :
    print("다운로드 날짜")
    for i in obj_dates:
        print(f"\t{i} ({df.get_weekday_from_yymmdd(i)})")
    check = input("계속 진행하려면 아무키나 입력하세요.")

# Chrome WebDriver를 실행한다.
#from webdriver_manager.chrome import ChromeDriverManager
#driver = webdriver.Chrome(ChromeDriverManager().install())

driver = webdriver.Chrome()
# 해당 URL로 이동한다.
grocery_url = 'http://134.209.127.114/invoice/inv_contractor_main/4/'
while True:
    ## 주소 접속할때까지 try
    try:
        driver.get(grocery_url)
        ## 로그인
        id = "contractor1"
        pw = "raincoming47"

        driver.find_element(By.ID, 'userName').send_keys(id)
        driver.find_element(By.ID, 'password').send_keys(pw)
        break
    except :
        print("재시도..")



login_button = driver.find_element(By.CSS_SELECTOR, '.btn.mt-3')
login_button.click()
## 로그인 성공
pages = []
stores = []
departs = []

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'card-header')))
for obj_date in obj_dates:
    pages, stores, departs = df.getDownloadUrl2(pages, stores, departs, driver,"grocery",obj_date)

houseware_url = 'http://134.209.127.114/invoice/inv_contractor_main/5/'
driver.get(houseware_url)
time.sleep(3)
for obj_date in obj_dates:
    pages, stores, departs = df.getDownloadUrl2(pages, stores, departs, driver,"houseware",obj_date)

if not os.path.isdir(f'C:/Users/user/Downloads/{today_mmdd}'):
    os.mkdir(f'C:/Users/user/Downloads/{today_mmdd}')

list_comb = zip(pages, stores, departs)
list_file_url = []

for item_comb in list_comb:
    driver.get(item_comb[0])
    if not os.path.isdir(f'C:/Users/user/Downloads/{today_mmdd}/{item_comb[1]}'):
        os.mkdir(f'C:/Users/user/Downloads/{today_mmdd}/{item_comb[1]}')
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    trs = soup.find_all('tr', {})
    idx_tr = 0
    for tr in trs:
        if idx_tr > 0:
            td_tag = tr.find_next('td')
            if td_tag is not None:
                td_tag = td_tag.find_next('td')
                if td_tag is not None:
                    td_tag = td_tag.find_next('td')
                    if td_tag is not None:
                        a_tag = td_tag.find_next('a')
                        file_url = a_tag["href"]
                        list_file_url.append((file_url, item_comb[1], item_comb[2]))
        idx_tr += 1

driver.quit()

dict_result = {}
list_thread = []
cnt_running = 0
total_cnt = 0
processed_cnt = 0
idx_thread = 0
with open('download_log.txt', 'a') as log_f:
    for item_file_url in list_file_url:
        if item_file_url[0] not in list_past:
            list_current.append(item_file_url[0]+'\n')
            dict_result[str(cnt_running)] = {'seq': cnt_running, 'url': item_file_url[0], 'state': '100', 'try_cnt': 0}
            th = threading.Thread(target=df.download, args=(item_file_url[0], today_mmdd, item_file_url[1], dict_result[str(cnt_running)]))
            th.start()
            list_thread.append(th)
            cnt_running += 1

    print("다운로드 경로 수집 완료.", cnt_running, "개 파일에 대해서 다운로드를 시작합니다.")

    total_cnt = cnt_running
    processed_cnt = 0
    idx_thread = 0
    while cnt_running > 0:
        th = list_thread[idx_thread]
        if not th.is_alive() or dict_result[str(idx_thread)]['state'] > '900':
            if dict_result[str(idx_thread)]['state'] == '900':
                log_f.write(f'{dict_result[str(idx_thread)]["seq"]}>> complete\n')
            else:
                log_f.write(f'{dict_result[str(idx_thread)]["seq"]}>> error try_cnt: {dict_result[str(idx_thread)]["try_cnt"]}\n')
            th.join()

            list_thread.pop(idx_thread)
            cnt_running -= 1
            processed_cnt += 1
        else:
            idx_thread += 1

        if idx_thread >= len(list_thread):
            idx_thread = 0

        df.progress_bar(processed_cnt, total_cnt, title="Download", bar_length=40)

print()
flag_fail = False
for i in range(total_cnt):
    if dict_result[str(i)]['state'] != '900':
        print(dict_result[str(i)])
        flag_fail = True

if flag_fail:
    print('다운받지 못한 파일이 있습니다. 다운로드 폴더를 삭제하고 재시도 해주세요.')
    # 다운로드 폴더 삭제 로직 추가 예정
    exit(1)

print(f'\n이번에 받은 uri 목록:')
print(list_current)
if len(list_current) == 0:
    import pyperclip
    kakaomsg = "오후 배분은 없습니다."
    pyperclip.copy(kakaomsg)
    print(f'단톡방 공지용 "{kakaomsg}" 가 클립보드에 저장되었습니다.')

fobj_dupchk = open(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}/{today_yymmdd}.txt', 'a')
fobj_dupchk.writelines(list_current)
fobj_dupchk.close()
fobj_dupchk = open(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}/{today_yymmdd}_last.txt', 'w')
fobj_dupchk.writelines(list_current)
fobj_dupchk.close()

if len(list_current) == 0:
    exit(1)

folder_name = today_mmdd + "/"
tmp_folder_name = today_mmdd + "_tmp" + "/"
year_str = today_yymmdd[0:2]
download_todayfolder= download_dir + folder_name
desktop_todayfolder = "C:/Users/user/Desktop/" + folder_name
desktop_todaytempfolder = "C:/Users/user/Desktop/" + tmp_folder_name
idx_upload = 0
uploadpath="C:/Users/user/Documents/GitHub/webdataentry/upload_data/" + today_yymmdd + "/"

while os.path.isdir(f'{uploadpath}'):
    uploadpath = "C:/Users/user/Documents/GitHub/webdataentry/upload_data/" + today_yymmdd + "_" + str(idx_upload) + "/"
    idx_upload += 1

cnt_moved_file = 0

## 바탕화면 오늘날짜 폴더 만들기
if not os.path.exists(desktop_todayfolder):
    df.makedir(desktop_todayfolder)
if not os.path.exists(desktop_todaytempfolder):
    df.makedir(desktop_todaytempfolder)
##
for store_name in os.listdir(download_todayfolder):
    ## 빈 폴더면 pass하기
    extracted_files = os.listdir(download_todayfolder)
    if (len(extracted_files) == 0):
        pass
    else:
        if not os.path.exists(desktop_todaytempfolder + store_name):
            df.makedir(desktop_todaytempfolder + store_name)
            print("폴더생성 :", store_name)
        for inv_file_fullname in os.listdir(download_todayfolder + store_name): ## 다운로드한 폴더에서 파일이동
            # if desktop_todayfolder + store_name + "/" + inv_file_fullname in list_past:
            #     continue
            # list_current.append(i+'\n')
            cnt_moved_file += 1
            shutil.move(download_todayfolder + store_name + "/" + inv_file_fullname, desktop_todaytempfolder + store_name + "/" + inv_file_fullname)
        os.rmdir(download_todayfolder + store_name)

#pdf파일 아닌것 제외하기
non_pdf_count=0
each_filenumber={}
print("\n")
#pdf 외 다른파일 있는지 검사
for i in os.listdir(desktop_todaytempfolder): ## 폴더이름
    non_pdf_count = 0
    for j in os.listdir(desktop_todaytempfolder+"/"+i): ##pdf파일이름
        if os.path.splitext(j)[1] in [".pdf", ".PDF"]:
            pass
        else:
            print(j + " ---> " + i)
            try:
                df.image_to_pdf(desktop_todaytempfolder + "/" + i + "/" + j)
                print("pdf로 변환완료!")
            except:
                print("pdf 변환실패 :", i, j)
                non_pdf_count += 1
        each_filenumber[i]= len(os.listdir(desktop_todaytempfolder+"/"+i)) - non_pdf_count

#upload 폴더에도 복사하기
if cnt_moved_file > 0:
    shutil.copytree(desktop_todaytempfolder, uploadpath)
    copy_tree(desktop_todaytempfolder, desktop_todayfolder)
shutil.rmtree(desktop_todaytempfolder)

print("전체 파일 :", cnt_moved_file)
print("전체 PDF 파일 :",sum(each_filenumber.values()))
print("DONE")
print("총 다운로드 zip파일 개수 :",len(list_current))
print("오늘 날짜는",df.getAbsTodayYYMMDD(),"입니다.")

if df.getAbsTodayYYMMDD() != today_yymmdd:
    print(today_yymmdd,"폴더로 작업했습니다. 재확인 바랍니다.")