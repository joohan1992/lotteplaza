import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import dataEntryFunction as df
import shutil
import time


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
print(list_past)

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
driver = webdriver.Chrome()
# 해당 URL로 이동한다.

grocery_url = 'http://134.209.127.114/invoice/inv_contractor_main/4/'
driver.get(grocery_url)

## 로그인
id = "contractor1"
pw = "raincoming47"

driver.find_element(By.ID, 'userName').send_keys(id)
driver.find_element(By.ID, 'password').send_keys(pw)

login_button = driver.find_element(By.CSS_SELECTOR, '.btn.mt-3')
login_button.click()
## 로그인 성공
download_urls = []
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'card-header')))
for obj_date in obj_dates:
    download_urls = df.getDownloadUrl(download_urls, driver,"grocery",obj_date)

houseware_url = 'http://134.209.127.114/invoice/inv_contractor_main/5/'
driver.get(houseware_url)
time.sleep(3)
for obj_date in obj_dates:
    download_urls = df.getDownloadUrl(download_urls, driver,"houseware",obj_date)



print("다운로드 경로 수집 완료.",len(download_urls),"개 파일에 대해서 다운로드를 시작합니다.")

for i in download_urls:
    driver.execute_script("window.open('" + i + "', '_blank');")



# 현재 시각 기록 - 파일이 생성된 이후의 시간으로 필터링
current_time = time.time()

while True:
    time.sleep(1) ## 1초마다 체크
    new_file_count = 0
    for filename in os.listdir(download_dir):
        full_path = os.path.join(download_dir, filename)
        if os.path.isfile(full_path):
            try:
                modification_time = os.path.getmtime(full_path)
                if modification_time >= current_time and (not full_path.endswith('.crdownload')):
                    new_file_count += 1
            except Exception:
                # TODO
                print('중간에 문제 발생 시')
    df.progress_bar(new_file_count, len(download_urls), title="Download", bar_length=40)

    if new_file_count == len(download_urls):
        print()
        print("다운로드가 완료되었습니다.")
        print("크롬을 종료합니다.")
        driver.quit()
        break

print(f'이번에 받은 uri 목록:')
print(list_current)
fobj_dupchk = open(f'C:/Users/user/Documents/GitHub/webdataentry/{dup_check_dir_nm}/{today_yymmdd}.txt', 'a')
fobj_dupchk.writelines(list_current)
fobj_dupchk.close()

# for i in obj_dates:
#     print("obj_date에",i,"가 있습니다.")
#     print("오늘 날짜로 설정된 날짜는", now_yymmdd,"입니다")
#     check = input("계속 진행하려면 아무키나 입력하세요")


for filename in os.listdir(download_dir):
    full_path = os.path.join(download_dir, filename)
    if os.path.isfile(full_path):
        modification_time = os.path.getmtime(full_path)
        if modification_time >= current_time and (not full_path.endswith('.crdownload')):
            df.makedir("C:/Users/user/Downloads/"+today_mmdd)
            shutil.move("C:/Users/user/Downloads/"+ filename, "C:/Users/user/Downloads/"+today_mmdd+"/" + filename)

## 다운로드 완료
## 다운로드 없이 압축만 풀고싶으면(수동다운로드 받았다면) 29~99번째 줄 주석처리하기


folder_name = today_mmdd + "/"
year_str = today_yymmdd[0:2]
download_todayfolder= download_dir + folder_name
desktop_todayfolder = "C:/Users/user/Desktop/" + folder_name
uploadpath="C:/Users/user/Documents/GitHub/webdataentry/upload_data/" + today_yymmdd + "/"
cnt_moved_file = 0
## 바탕화면 오늘날짜 폴더 만들기
if not os.path.exists(desktop_todayfolder):
    df.makedir(desktop_todayfolder)
##
for i in os.listdir(download_todayfolder):
    print("\t압축해제 :", i)
    zipf = zipfile.ZipFile(download_todayfolder + i)
    zipf.extractall(download_todayfolder + os.path.splitext(i)[0])
    zipf.close()
    store_name = i.split("_")[1] ## VA900 같은 store name
    zipfile_name = os.path.splitext(i)[0]
    extracted_folder = os.path.splitext(i)[0] ## 폴더 이름

    ## 빈 압축파일이면 pass하기
    extracted_files = zipf.namelist()

    if (len(extracted_files) == 0):
        pass

    else:
        if not os.path.exists(desktop_todayfolder + store_name):
            df.makedir(desktop_todayfolder + store_name)
            print("폴더생성 :", store_name)
        for j in os.listdir(download_todayfolder + extracted_folder): ## 압축 해제된 폴더에서 파일이동
            # if desktop_todayfolder + store_name + "/" + inv_file_fullname in list_past:
            #     continue
            # list_current.append(i+'\n')
            cnt_moved_file += 1
            inv_file_fullname = j
            inv_file_name, inv_file_ext = os.path.splitext(j)[0], os.path.splitext(j)[1]
            uniq = 1
            while os.path.exists(desktop_todayfolder + store_name + "/" + inv_file_fullname):  # 동일한 파일명이 존재할 때
                inv_file_fullname = inv_file_name + "_(" + str(uniq) + ")." + inv_file_ext  # 파일명(1) 파일명(2)...
                uniq += 1
            shutil.move(download_todayfolder + zipfile_name + "/" + j, desktop_todayfolder + store_name + "/" + inv_file_fullname)
        os.rmdir(download_todayfolder + extracted_folder)

#upload 폴더에도 복사하기
shutil.copytree(desktop_todayfolder, uploadpath)

#pdf파일 아닌것 제외하기
non_pdf_count=0
each_filenumber={}
print("\n")
#pdf 외 다른파일 있는지 검사
for i in os.listdir(uploadpath): ## 폴더이름
    non_pdf_count = 0
    for j in os.listdir(uploadpath+"/"+i): ##pdf파일이름
        if os.path.splitext(j)[1] in [".pdf", ".PDF"]:
            pass
        else:
            print(j+" ---> "+i)
            non_pdf_count += 1
        each_filenumber[i]= len(os.listdir(uploadpath+"/"+i)) - non_pdf_count

print("전체 파일 :", cnt_moved_file)
print("전체 PDF 파일 :",sum(each_filenumber.values()))
print("DONE")
print("총 다운로드 zip파일 개수 :",len(download_urls))
print("오늘 날짜는",df.getAbsTodayYYMMDD(),"입니다.")

if df.getAbsTodayYYMMDD() != today_yymmdd:
    print(today_yymmdd,"폴더로 작업했습니다. 재확인 바랍니다.")