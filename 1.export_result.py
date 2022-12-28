import requests
import os
import time
import shutil
import zipfile
from glob import glob
import datetime
def makedir(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print("폴더생성 :",directory)
    except OSError:
        print("Error: Failed to create the directory.")
def setfoldername1():








#########################################################################################################
#########################################################################################################


#                     *  1,2 하나로 합쳤습니다.
#                     *  다운받고 > 압축풀고 > 분류하는거까지 실행하는 코드입니다.

#########################################################################################################
#########################################################################################################
#########################################################################################################
#########################################################################################################
                                                                                                #########
                                                                                                #########
                    filename = "TOT_221227_01"  ## 가운데 오늘날짜 쓰기     ex) TOT_221221_01    #########
                    objdate  = ""                ## -1로 할거면 ""로 두기   ex) 20221215         #########
                                                                                                #########
                                                                                                #########
#########################################################################################################
#########################################################################################################
#########################################################################################################
#########################################################################################################














                    today = filename.split("_")[1][2:]
                    return today, filename, objdate



dir = "C:/Users/user/Documents/GitHub/lotteplaza/작업파일/"
gitdir="C:/Users/user/Documents/GitHub/webdataentry/result/"
today,filename,objdate = setfoldername1()
## 다운로드
url=f"http://10.28.78.30:8889/export_result/-1/{filename}/{objdate}/-1"
if objdate == "":
    url = f"http://10.28.78.30:8889/export_result/-1/{filename}/-1/-1"
else:
    url = f"http://10.28.78.30:8889/export_result/-1/{filename}/{objdate}/-1"
print(url)
response = requests.get(url)
print("\tresponse :", response.status_code)
time.sleep(2)
makedir(dir + today)
makedir(dir + today+"/제외")
makedir(dir + today+"/3차")
makedir(dir + today+"/4차")
makedir(dir + today+"/결과")
makedir(dir + today+"/결과/MARGIN_DIFF")
shutil.copy2(gitdir + filename + "_except.zip",
            dir + today + "/" + filename + "_except.zip")
shutil.copy2(gitdir + filename + "_result.xlsx",
            dir + today+"/3차" + "/" + filename + "_result.xlsx")

zipf = zipfile.ZipFile(dir + today + "/" + filename + "_except.zip")
zipf.extractall(dir + today + "/제외")
zipf.close()
print("다운로드 및 파일 이동 완료.")
print("제외 파일 분류를 시작합니다.")

## 제외 파일 분류
path=f"C:/Users/user/Documents/GitHub/lotteplaza/작업파일/{today}/제외/"
dirpath=f"C:/Users/user/Documents/GitHub/lotteplaza/작업파일/{today}/제외/*.pdf"
dict={"업체코드 검색 불가" : [], "입력 대상 제품 없음" : [], "제품코드, UPC 부재 & 한글 Description & 수기" : [], "파본" : []}
totalnum=len(glob(dirpath))
print("==============================")
print("작업 폴더 :", today)
print("==============================")
print("전체 제외 파일 :", totalnum)
for x in glob(dirpath):
    fname=os.path.basename(x)
    rst = fname.translate(str.maketrans('1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ', '..........abcdefghijklmnopqrstuvwxyz')).split(".")[-2]
    if "제외" in rst and "업체" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "제외" in rst and "업종" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "내용" in rst and "없음" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "rheebros" in rst or "sungwon" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "패킹" in rst or "단가" in rst or "매장이동" in rst or "크레딧" in rst or "크래딧" in rst or "credit" in rst or "이동" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "리턴" in rst or "return" in rst or "단가" in rst or "charge" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "statement" in rst or "집계" in rst or "stock" in rst or "receip" in rst or "중복" in rst or "리스트" in rst:
        dict["입력 대상 제품 없음"].append(fname)
    elif "코드없음" in rst or "코드 없음" in rst:
        dict["제품코드, UPC 부재 & 한글 Description & 수기"].append(fname)
    elif "수기" in rst or "제품번호" in rst:
        dict["제품코드, UPC 부재 & 한글 Description & 수기"].append(fname)
    elif "upc" in rst or "한글" in rst:
        dict["제품코드, UPC 부재 & 한글 Description & 수기"].append(fname)
    elif "파본" in rst:
        dict["파본"].append(fname)
    elif "업체코드" in rst:
        dict["업체코드 검색 불가"].append(fname)
    else:
        print(fname)
print("------------------------------")
print("분류 완료, 파일 이동...")
filemovenum=0
for i in dict:
    if len(dict[i])>0:
        makedir(path+i)
        print("\t",i,":",len(dict[i]))
        filemovenum+=len(dict[i])
    for k in dict[i]:
        shutil.move(os.path.join(path,k), os.path.join(path+i,k))
print("------------------------------")
print(filemovenum,"개 파일 이동 완료")

print("제외 폴더 압축 시작")
zip_path=f"C:/Users/user/Documents/GitHub/lotteplaza/작업파일/{today}/제외"
result_path=f"C:/Users/user/Documents/GitHub/lotteplaza/작업파일/{today}/결과/제외인보이스_{today+str(datetime.datetime.now().year)}"
shutil.make_archive(result_path,'zip',zip_path)
print("제외 폴더 압축 완료")

if filemovenum==totalnum:
    print("Done")
else:
    print("전체 제외 파일과 이동된 파일의 개수가 다릅니다.")
    print("수동 확인이 필요합니다.")


today = datetime.datetime.now()
year=str(today.year)[2:4]
day=str(today.day).zfill(2)
month=str(today.month).zfill(2)
if year+month+day != filename.split("_")[1] :
    print("오늘 날짜는",year+month+day,"입니다.")
    print(filename.split("_")[1],"폴더로 작업했습니다. 재확인 바랍니다.")