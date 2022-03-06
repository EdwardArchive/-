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
    #print(data.isna().sum())
    print(name+"고객님의 "+infomiss+" 정보가 비어있습니다.")
    return "{infomiss}정보가 비어있습니다.".format(infomiss=infomiss)

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

def process(name,file_path,case,sheet_name):
    #2. 파일복사
    if case ==  "컨설팅":
        shutil.copyfile(r"{file_path}".format(file_path=file_path), r"{file_path}".format(file_path=file_path).replace("(장기)","({name})".format(name=name)).replace("/컨설팅 신청서/","/컨설팅 신청서/중간결과물/"))
    elif case == "계획서":
        shutil.copyfile(r"{file_path}".format(file_path=file_path), r"{file_path}".format(file_path=file_path).replace("양식"," {name}".format(name=name)).replace("/전직계획서/","/전직계획서/중간결과물/"))
    else :
        shutil.copyfile(r"{file_path}".format(file_path=file_path), r"{file_path}".format(file_path=file_path).replace("양식","_{name}".format(name=name)).replace("/발송주소/","/발송주소/중간결과물/"))
    
    #3. 한글열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    #4. 보안모듈 적용
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

    #5. 복사한 한글 열기
    if case == "컨설팅" :
        hwp.Open(r"{file_path}".format(file_path=file_path).replace("(장기)","({name})".format(name=name)).replace("/컨설팅 신청서/","/컨설팅 신청서/중간결과물/"))
    elif case == "계획서":
        hwp.Open(r"{file_path}".format(file_path=file_path).replace("양식"," {name}".format(name=name)).replace("/전직계획서/","/전직계획서/중간결과물/"))
    else :
        hwp.Open(r"{file_path}".format(file_path=file_path).replace("양식","_{name}".format(name=name)).replace("/발송주소/","/발송주소/중간결과물/"))

    #6. 한글 누름틀 목록 불러오기
    field_list = [i for i in hwp.GetFieldList().split("\x02")]

    #7. 데이터 처리
    y = df.loc[ df [ '성명' ] == name ]
    print("\n"+name+"고객님의 {case} 자료를 생성중입니다...".format(case=case))
    try:
        birth = ''.join(y['생년월일'])
        birth = datetime.strptime(birth,'%Y-%m-%d').date()
        today = datetime.now().date()
        year = today.year - birth.year
        if today.month < birth.month or (today.month == birth.month and today.day < birth.day):
            year-=1
        info_dic=dict({
            '장기중기' : '☑ 장기' if '장기' in sheet_name else '☑ 중기',
            'N차' : sheet_name.split('-')[1][0:1],
            '성명' : ''.join(y['성명']),
            '생년월일' : ''.join(y['생년월일']),
            '나이' : year,
            '연락처' : ''.join(y['연락처']) if y['연락처'].isna().sum()==0  else alrt(name,'전화번호'),
            '이메일' : ''.join(y['E-mail 주소']) if y['E-mail 주소'].isna().sum()==0  else alrt(name,'이메일'),
            '기본주소' : ''.join(y['기본주소']) if y['기본주소'].isna().sum()==0  else alrt(name,'기본주소'),
            '군' : ''.join(y['군']) if y['군'].isna().sum()==0  else alrt(name,'군'),
            '계급' : ''.join(y['계급']) if y['계급'].isna().sum()==0  else alrt(name,'계급'),
            '병과' : ''.join(y['병과']) if y['병과'].isna().sum()==0  else alrt(name,'병과'),
            '군번' : ''.join(y['군번']) if y['군번'].isna().sum()==0  else alrt(name,'군번'),
            '전역예정일' : ''.join(y['전역(예정)일']) if y['전역(예정)일'].isna().sum()==0  else alrt(name,'전역예정일'),
            '복무기간_년' : ''.join(y['복무기간']).split(" ")[0].replace("년",'') if y['복무기간'].isna().sum()==0  else alrt(name,'복무기간_년'),
            '복무기간_월' : ''.join(y['복무기간']).split(" ")[1].replace("개월",'') if y['복무기간'].isna().sum()==0  else alrt(name,'복무기간_월'),
            '권역_영등포' : '√' if '영등포' == ''.join(y['센터']) else '',
            '권역_서울숲' : '√' if '서울' == ''.join(y['센터']) else '',
            '권역_수원' : '√' if '수원' == ''.join(y['센터']) else '',
            '권역_파주' : '√' if '파주' == ''.join(y['센터']) else '',
            '동의함' : '☑',
            '수신자' : ''.join(y['성명']) + " 선생님",
            '우편번호' : '00000'
        })
        #8. 필드에 데이터 입력
        for field in field_list:
            hwp.PutFieldText(f'{field}', info_dic[field])

        hwp.Save(False)
        hwpkill()
        print("\n"+name+"고객님의 {case} 자료를 생성완료했습니다!".format(case=case))

    except Exception as e:
        print("\n"+name+"고객님의 {case} 생성실패했습니다! 에러코드 {e}".format(case=case,e=e))
        hwpkill()

if __name__ == "__main__":
    file_path = ""
    file_path_2 = ""
    file_path_3 = ""
    file_path_ex = ""

    #1. 문서 위치 설정
    root = tk.Tk()
    root.withdraw()

    option = str(input('원하시는 옵션은 무엇입니까?(1. 컨설팅 + 계획서, 2.컨설팅 + 계획서 + 발송주소) : '))
    if option == "2" or option =="1":
        messagebox.showinfo("사용법", "컨설팅 신청서를 선택하여 주세요!")
        file_path = filedialog.askopenfilename(filetypes=[("한글 문서(hwp)", "*.hwp")])

        messagebox.showinfo("사용법", "전직 계획서를 선택하여 주세요!")
        file_path_2 = filedialog.askopenfilename(filetypes=[("한글 문서(hwp)", "*.hwp")])
        
        if option =="2":
            messagebox.showinfo("사용법", "발송주소 양식을 선택하여 주세요!")
            file_path_3 = filedialog.askopenfilename(filetypes=[("한글 문서(hwp)", "*.hwp")])

        messagebox.showinfo("사용법", "엑셀을 선택하여 주세요!")
        file_path_ex = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        print(file_path)
        print(file_path_ex)

    answer = str(input('원하시는 사람 혹은 옵션은 무엇입니까?(사람이름,전부) : '))
    sheet_name = str(input('sheet 이름은 무엇입니까? : '))
    print(answer)
    #1. 엑셀파일 열기
    df = pd.read_excel(r'{}'.format(file_path_ex), sheet_name=sheet_name,header=2)

    if(answer=='전부'): 
        for name in df['성명'].tolist():
            process(name,file_path,"컨설팅",sheet_name)
            process(name,file_path_2,"계획서",sheet_name)
            if option == "2" : process(name,file_path_3,"주소양식",sheet_name)
    else :
        for name in answer.split(','):
            process(name,file_path,"컨설팅",sheet_name)
            process(name,file_path_2,"계획서",sheet_name)
            if option == "2" : process(name,file_path_3,"주소양식",sheet_name)

        


        