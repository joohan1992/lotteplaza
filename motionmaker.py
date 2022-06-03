dir_org = './loading_img/org/'
dir_rename = './loading_img/rename/'
for i in range(144):
    org_file_nm = "loading_"+str(i).zfill(5)+".png"
    rename = "img_loading_"+str(i).zfill(5)+".png"
    org_file = open(dir_org+org_file_nm, 'rb')
    rename_file = open(dir_rename+rename, 'wb')
    rename_file.writelines(org_file.readlines())
    rename_file.close()
    org_file.close()
    print('<item android:drawable="@drawable/img_loading_'+str(i).zfill(5)+'.png" android:duration="10" />')