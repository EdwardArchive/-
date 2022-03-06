#pip install pandas
#pip install xlrd
#pip install pywin32
#pip install openpyxl


import pandas as pd
import shutil
import win32com.client as win32
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import psutil

def alrt(name,infomiss,data=0):
    #print(data.isna()[1])
    print(name+"고객님의 "+infomiss+" 정보가 비어있습니다.")
    return "{infomiss} 가 비어있습니다.".format(infomiss=infomiss)

def hwpkill():
    for proc in psutil.process_iter():
        try:
            # 프로세스 이름, PID값 가져오기
            processName = proc.name()
            processID = proc.pid

            if processName == "Hwp.exe":
                parent_pid = processID  #PID
                parent = psutil.Process(parent_pid) # PID 찾기

                for child in parent.children(recursive=True):  #자식-부모 종료
                    child.kill()

                parent.kill()
 
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    root.quit()

def process(name):
    #2. 파일복사
    shutil.copyfile(r"{file_path}".format(file_path=file_path), r"{file_path}".format(file_path=file_path).replace("(장기)","({name})".format(name=name)))
    
    #3. 한글열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    #4. 보안모듈 적용
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

    #5. 복사한 한글 열기
    hwp.Open(r"{file_path}".format(file_path=file_path).replace("(장기)","({name})".format(name=name)))

    #6. 한글 누름틀 목록 불러오기
    field_list = [i for i in hwp.GetFieldList().split("\x02")]

    #7. 데이터 처리
    y = df.loc [ df [ '성명' ] == name ]
    print("\n"+name+"고객님의 자료를 생성중입니다...")
    try:
        birth = y['생년월일'].astype(str)[0]
        birth = datetime.strptime(birth,'%Y-%m-%d').date()
        today = datetime.now().date()
        year = today.year - birth.year
        if today.month < birth.month or (today.month == birth.month and today.day < birth.day):
            year-=1
        info_dic=dict({
            '장기' : '☑',
            '성명' : y['성명'].astype(str)[0],
            '생년월일' : y['생년월일'].astype(str)[0],
            '나이' : year,
            '전화번호' : y['연락처'].astype(str)[0],
            '이메일' : y['E-mail 주소'].astype(str)[0],
            '기본주소' : y['기본주소'].astype(str)[0],
            '군' : y['군'].astype(str)[0],
            '계급' : y['계급'].astype(str)[0],
            '병과' : y['병과'].astype(str)[0],
            '군번' : y['군번'].astype(str)[0],
            '전역예정일' : y['전역(예정)일'].astype(str)[0],
            '복무기간_년' : y['복무기간'].astype(str)[0].split(" ")[0].replace("년",''),
            '복무기간_월' : y['복무기간'].astype(str)[0].split(" ")[1].replace("개월",''),
            '권역_영등포' : '√' if '영등포' == y['센터'].astype(str)[0] else '',
            '권역_서울숲' : '√' if '서울' == y['센터'].astype(str)[0] else '',
            '권역_수원' : '√' if '수원' == y['센터'].astype(str)[0] else '',
            '권역_파주' : '√' if '파주' == y['센터'].astype(str)[0] else ''
        })
    except :
        birth = ''.join(y['생년월일'])
        birth = datetime.strptime(birth,'%Y-%m-%d').date()
        today = datetime.now().date()
        year = today.year - birth.year
        if today.month < birth.month or (today.month == birth.month and today.day < birth.day):
            year-=1
        info_dic=dict({
            '장기' : '☑',
            '성명' : ''.join(y['성명']),
            '생년월일' : ''.join(y['생년월일']),
            '나이' : year,
            '전화번호' : ''.join(y['연락처']) if y['연락처'].isna()[1]==True  else alrt(name,'전화번호'),
            '이메일' : ''.join(y['E-mail 주소']) if y['E-mail 주소'].isna()[1]==True  else alrt(name,'이메일'),
            '기본주소' : ''.join(y['기본주소']) if y['기본주소'].isna()[1]==True  else alrt(name,'기본주소'),
            '군' : ''.join(y['군']) if y['E-mail 주소'].isna()[1]==True  else alrt(name,'군'),
            '계급' : ''.join(y['계급']) if y['계급'].isna()[1]==True  else alrt(name,'계급'),
            '병과' : ''.join(y['병과']) if y['병과'].isna()[1]==True  else alrt(name,'병과'),
            '군번' : ''.join(y['군번']) if y['군번'].isna()[1]==True  else alrt(name,'군번'),
            '전역예정일' : ''.join(y['전역(예정)일']) if y['E-mail 주소'].isna()[1]==True  else alrt(name,'전역예정일'),
            '복무기간_년' : ''.join(y['복무기간']).split(" ")[0].replace("년",'') if y['복무기간'].isna()[1]==True  else alrt(name,'복무기간_년'),
            '복무기간_월' : ''.join(y['복무기간']).split(" ")[1].replace("개월",'') if y['복무기간'].isna()[1]==True  else alrt(name,'복무기간_월'),
            '권역_영등포' : '√' if '영등포' == ''.join(y['센터']) else '',
            '권역_서울숲' : '√' if '서울' == ''.join(y['센터']) else '',
            '권역_수원' : '√' if '수원' == ''.join(y['센터']) else '',
            '권역_파주' : '√' if '파주' == ''.join(y['센터']) else ''
        })
       
    #8. 필드에 데이터 입력
    for field in field_list:
        hwp.PutFieldText(f'{field}', info_dic[field])

    hwp.Save(False)
    hwpkill()
    print("\n"+name+"고객님의 자료를 생성완료했습니다!")
#1. 문서 위치 설정
root = tk.Tk()
root.withdraw()

messagebox.showinfo("사용법", "한글문서를 선택하여 주세요!")
file_path = filedialog.askopenfilename(filetypes=[("한글 문서(hwp)", "*.hwp")])

messagebox.showinfo("사용법", "엑셀을 선택하여 주세요!")
file_path_ex = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
print(file_path)
print(file_path_ex)

answer = str(input('원하시는 사람 혹은 옵션은 무엇입니까?(사람이름,전부) : '))
sheet_name = str(input('sheet 이름은 무엇입니까? : '))

#1. 엑셀파일 열기
df = pd.read_excel(r'{}'.format(file_path_ex), sheet_name=sheet_name,header=2)

if(answer=='전부'): 
    for name in df['성명'].tolist():
         process(name)
else : process(answer)

    


        