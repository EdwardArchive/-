from distutils.log import info
from tkinter.tix import Tree
from matplotlib.pyplot import getp
from numpy import empty
import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime
from getpass import getpass
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# 엑셀데이터 알람
def alrt(name,infomiss,data=0):
    #print(data.isna().sum())
    print(name+"고객님의 "+infomiss+" 정보가 비어있습니다.")
    return "{infomiss}정보가 비어있습니다.".format(infomiss=infomiss)

# 데이터 처리
def process(name,sheet_name):
    info_dic = dict()
    y = df.loc[ df [ '성명' ] == name ]
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
            '생년월일' : ''.join(y['생년월일']) if y['생년월일'].isna().sum()==0  else alrt(name,'생년월일'),
            '성별' : ''.join(y['성별']) if y['성별'].isna().sum()==0  else alrt(name,'성별'),
            '소속' : ''.join(y['소속']) if y['소속'].isna().sum()==0  else alrt(name,'소속'),
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

    
    except Exception as e:
        print("\n"+name+"고객님의 엑셀 정보획드에 실패하였습니다 에러코드 {e}".format(e=e))
        
    return info_dic

messagebox.showinfo("사용법", "엑셀을 선택하여 주세요!")
file_path_ex = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
sheet_name = str(input('sheet 이름은 무엇입니까? : '))
id = str(input('아이디를 입력하세요 : '))
pwd = getpass('비밀번호를 입력하세요 : ')
service_start = str(input('서비스 시작일자 : '))
service_end = str(input('서비스 종료일자 : '))
service_add = str(input('서비스 등록일자 : '))
while True:
    
    
    #1. 엑셀파일 열기
    df = pd.read_excel(r'{}'.format(file_path_ex), sheet_name=sheet_name,header=2)
    print("\n----엑셀파일에 있는 이름들입니다----")
    print(df['성명'])
    print("----------------------------------\n")
    while True:
        name = str(input('처리하고자 하는 사람의 이름을 입력하세요 : '))
        if df['성명'].str.contains(name).any() :
            break
        else:
            print("이름을 다시 입력하세요 잘못입력했거나 excel 파일에 해당 이름이 없습니다.")

    #2. 사이트 접속
    URL = 'https://cloud.sscrm.co.kr/Login.aspx'

    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver = webdriver.Chrome(executable_path='chromedriver',options=options)
    driver.get(url=URL)

    #3. 로그인작업
    driver.find_element_by_xpath('//*[@id="txtLoginID"]').send_keys('{id}'.format(id=id))
    driver.find_element_by_xpath('//*[@id="txtPassword"]').send_keys('{pwd}'.format(pwd=pwd))
    driver.find_element_by_xpath('//*[@id="btnLogin"]').click()

    #4. 고객 정보 입력 이동
    driver.get(url="https://cloud.sscrm.co.kr/Pages__C002/CS/CSCI_List.aspx")
    driver.find_element_by_xpath('//*[@id="ucLeft_listCS12"]/a').click()

    #5. 고객 정보 입력 
    get_info_dict = process(name,sheet_name)
    if get_info_dict is not empty: 
        driver.find_element_by_xpath("//select[@name='ctl00$ContentPage$FLD_CTCCODE']/option[text()='파주지사']").click() # CTC
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_CUSTNM"]').send_keys(get_info_dict['성명']) # 고객명
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_EMAIL"]').send_keys(get_info_dict['이메일']) # 이메일
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_MOBILE"]').send_keys(get_info_dict['연락처']) # 휴대폰
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_BIRTHDAY"]').send_keys(get_info_dict['생년월일']) # 생년월일
        driver.find_element_by_xpath("//select[@name='ctl00$ContentPage$FLD_SEX']/option[text()='{sex}']".format(sex=get_info_dict['성별'])).click() # 성별
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_WORKCOMPANYNAME"]').send_keys(get_info_dict['군'])
        driver.find_element_by_xpath("//select[@name='ctl00$ContentPage$FLD_CAREERGOALCODE']/option[text()='취업']").click()  # 경력목표
        driver.find_element_by_xpath("//select[@name='ctl00$ContentPage$FLD_PROCESSSTATUSCODE']/option[text()='진행중']").click() # 진행상황
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_SERVICEBEGINDATE"]').send_keys(service_start) # 서비스 시작일자
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_SERVICEENDDATE"]').send_keys(service_end) # 서비스 종료일자
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_SERVICEREGDATE"]').send_keys(service_add) # 서비스 등록일자
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_ADDRESS1"]').send_keys('{addr}'.format(addr=get_info_dict['기본주소'])) # 주소
        if get_info_dict['군'] =='해병대':
            driver.find_element_by_xpath('//*[@id="ContentPage_FLD_WORKRANK"]').send_keys('{army} {rank}'.format(army=get_info_dict['군'][:2],rank=get_info_dict['계급'])) # 직급계급
        else :
            driver.find_element_by_xpath('//*[@id="ContentPage_FLD_WORKRANK"]').send_keys('{army} {rank}'.format(army=get_info_dict['군'][:1],rank=get_info_dict['계급'])) # 직급계급
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_WORKRETIREDATE"]').send_keys('{}'.format(get_info_dict['전역예정일'])) # 퇴역일 
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_ARMYWORKTYPE"]').send_keys('{armytype}'.format(armytype=get_info_dict['병과'])) # 병과
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_WORKDEPARTMENT"]').send_keys('{armytype}'.format(armytype=get_info_dict['소속'])) # 부서
        driver.find_element_by_xpath('//*[@id="ContentPage_FLD_WORKYEAR"]').send_keys('{}년'.format(get_info_dict['복무기간_년'])) # 근무기간 년
        if get_info_dict['복무기간_월']!='0':
            driver.find_element_by_xpath("//select[@name='ctl00$ContentPage$FLD_WORKMONTH']/option[text()='{WORKMONTH}']".format(WORKMONTH=get_info_dict['복무기간_월'])).click()  # 근무기간 월
        driver.find_element_by_xpath("//select[@name='ctl00$ContentPage$FLD_USEYN']/option[text()='사용가능']").click() # 커리어플러스사용여부

    else :
        messagebox.showinfo("사용법", "엑셀 데이터를 확인해주세요")
        driver.close()

    if input("종료하시겠습니까? (Y/N) : ") == 'Y':
        break   