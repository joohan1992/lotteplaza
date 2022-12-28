import os
import zipfile
import shutil
import datetime

src = '/home/banana/'
dir = '/home/banana/txt/'
d_path="C:/Users/user/Downloads/"
def setfoldername():
            """





    1. zip파일 다운로드
    2. [다운로드]폴더 하위에 오늘날짜로 폴더생성 (ex 1209)후 다운로드받은 zip파일 이동
    3. 아래 folder_name에 오늘날짜 4자리 입력
    4. 실행



"""
#######################################################
#######################################################
                                              #########
                                              #########
            folder_name = "1228"              #########
            year_str = "22"                   #########
                                              #########
                                              #########
#######################################################
#######################################################










            return str(folder_name)+"/", str(year_str)
folder_name, year_str = setfoldername()
path=d_path+folder_name
homepath="C:/Users/user/Desktop/"+folder_name
uploadpath="C:/Users/user/Documents/GitHub/webdataentry/upload_data/"+year_str+folder_name

if not os.path.exists(homepath):
    os.mkdir(homepath)
for i in os.listdir(path):
    zipf = zipfile.ZipFile(path+i)
    zipf.extractall(path+i.split(".zip")[0])
    zipf.close()
    if not os.path.exists(homepath+i.split("_")[1]):
        os.mkdir(homepath+i.split("_")[1])
    for j in os.listdir(path+i.split(".zip")[0]):
        shutil.move(path+i.split(".zip")[0]+"/"+j, homepath+i.split("_")[1] +"/"+j)
    os.rmdir(path+i.split(".zip")[0])
shutil.copytree(homepath, uploadpath)
for i in os.listdir(uploadpath):
    for j in os.listdir(uploadpath+"/"+i):
        if j.split(".")[-1] in ["pdf", "PDF"]:
            pass
        else:
            print(i+" ---> "+j)
print("DONE")

today = datetime.datetime.now()
year=str(today.year)[2:4]
day=str(today.day).zfill(2)
month=str(today.month).zfill(2)

if year+month+day != year_str+folder_name+"/":
    print("오늘 날짜는",year+month+day,"입니다.")
    print(year_str+folder_name,"폴더로 작업했습니다. 재확인 바랍니다.")

















































