# -*- coding: utf-8 -*-

import docxpy
import re
import pandas as pd
import glob
from natsort import natsorted
import pickle
from tqdm import tqdm
import csv 
import re

### File loading ###
path_dir = '/Users/chanhee.kang/Desktop/sa_paser/test2.docx' #한국
file_list = natsorted(glob.glob(path_dir)) #natural sorted
publisher = None

mwp_df = pd.read_csv(r"/Users/chanhee.kang/Desktop/sa_paser/test.csv")

curr_number = 0
country_list = []
## Body & Classification ##
docu_body_re = re.compile(r'(?<=Body\n\n\n\n\n\n)((.|\n)*)(?=Classification)')
docu_body_graphic_re = re.compile(r'(?<=Body\n\n\n\n\n\n)((.|\n)*)(?=Graphic)')
docu_body_load_date_re =re.compile(r'(?<=Body\n\n\n\n\n\n)((.|\n)*)(?=Load Date:)')
docu_classification = re.compile(r'(?<=Classification\n\n\n\n\n\n)((.|\n)*)(?=)')
docu_date = re.compile(r'((((((((Jan(uary)?)|(Mar(ch)?)|(May)|(July?)|(Aug(ust)?)|(Oct(ober)?)|(Dec(ember)?)) ((3[01])|29))|(((Apr(il)?)|(June?)|(Sep(tember)?)|(Nov(ember)?)) ((30)|(29)))|(((Jan(uary)?)|(Feb(ruary)?|(Mar(ch)?)|(Apr(il)?)|(May)|(June?)|(July?)|(Aug(ust)?)|(Sep(tember)?)|(Oct(ober)?)|(Nov(ember)?)|(Dec(ember)?))) (2[0-8]|(1\d)|(0?[1-9])))),? )|(((((1[02])|(0?[13578]))[\.\-/]((3[01])|29))|(((11)|(0?[469]))[\.\-/]((30)|(29)))|(((1[0-2])|(0?[1-9]))[\.\-/](2[0-8]|(1\d)|(0?[1-9]))))[\.\-/])|(((((3[01])|29)[ \-\./]((Jan(uary)?)|(Mar(ch)?)|(May)|(July?)|(Aug(ust)?)|(Oct(ober)?)|(Dec(ember)?)))|(((30)|(29))[ \.\-/]((Apr(il)?)|(June?)|(Sep(tember)?)|(Nov(ember)?)))|((2[0-8]|(1\d)|(0?[1-9]))[ \.\-/]((Jan(uary)?)|(Feb(ruary)?|(Mar(ch)?)|(Apr(il)?)|(May)|(June?)|(July?)|(Aug(ust)?)|(Sep(tember)?)|(Oct(ober)?)|(Nov(ember)?)|(Dec(ember)?)))))[ \-\./])|((((3[01])|29)((Jan)|(Mar)|(May)|(Jul)|(Aug)|(Oct)|(Dec)))|(((30)|(29))((Apr)|(Jun)|(Sep)|(Nov)))|((2[0-8]|(1\d)|(0[1-9]))((Jan)|(Feb)|(Mar)|(Apr)|(May)|(Jun)|(Jul)|(Aug)|(Sep)|(Oct)|(Nov)|(Dec)))))(((175[3-9])|(17[6-9]\d)|(1[89]\d{2})|[2-9]\d{3})|\d{2}))|((((175[3-9])|(17[6-9]\d)|(1[89]\d{2})|[2-9]\d{3})|\d{2})((((1[02])|(0[13578]))((3[01])|29))|(((11)|(0[469]))((30)|(29)))|(((1[0-2])|(0[1-9]))(2[0-8]|(1\d)|(0[1-9])))))|(((29Feb)|(29[ \.\-/]Feb(ruary)?[ \.\-/])|(Feb(ruary)? 29,? ?)|(0?2[\.\-/]29[\.\-/]))((((([2468][048])|([3579][26]))00)|(17((56)|([68][048])|([79][26])))|(((1[89])|([2-9]\d))(([2468][048])|([13579][26])|(0[48]))))|(([02468][048])|([13579][26]))))|(((((([2468][048])|([3579][26]))00)|(17((56)|([68][048])|([79][26])))|(((1[89])|([2-9]\d))(([2468][048])|([13579][26])|(0[48]))))|(([02468][048])|([13579][26])))(0229)))')
Master = pd.DataFrame(columns=['Publisher', 'Country', 'Title', 'Date', 'Body'])  # 빈 데이터프레임 만들기
idx = 0

##For loop, 지정한 파일 리스트에서 추출
for file in tqdm(file_list):
    #print(file)
    doc = docxpy.DOCReader(file)
    doc.process()
    doc.data['document'] = re.sub(r'[\?\.\|\*\[\]\$\+\-]',' ',doc.data['document'])  #물음표, 마침표 등 정규식과 겹치는 것 삭제해야함(괄호 제외)
    target_str = doc.data['document']
    max = 0

    for pub in mwp_df.values[:,0]:

        length = len([m.start() for m in re.finditer(pub, target_str)])
        if(0 == length):
            continue
        else:
            publisher = pub
        if(publisher == None):
            continue

        title_publisher = re.compile(r'.*(?=\n\n' + publisher + '\s*?\n)')  # KH

        # title_publisher_list = [x for x in title_publisher.findall(doc.data['document']) if x!='']
        title_publisher_list = [x for x in title_publisher.findall(doc.data['document']) if
                                (x != '') & (re.match(r'\S.*', x) != None)]
        for i in range(0, len(title_publisher_list)):
            # print(i)
            if (i != len(title_publisher_list) - 1):
                docu_whole_re = re.compile(
                    r'(?<=\n{})((.|\n)*)(?=\n({}))'.format(title_publisher_list[i], title_publisher_list[i + 1]))
            else:
                docu_whole_re = re.compile(r'(?<=\n{})((.|\n)*)(?=\n(End of Document))'.format(title_publisher_list[i]))
            if (docu_whole_re.search(doc.data['document']) != None):
                total_text = docu_whole_re.search(doc.data['document']).group()
                if (re.search(r'\n\nGraphic\n', total_text) != None):
                    if (docu_body_graphic_re.search(total_text) != None):
                        body_text = docu_body_graphic_re.search(total_text).group()
                    else:
                        body_text = ''
                else:
                    if (docu_body_re.search(total_text) != None):
                        body_text = docu_body_re.search(total_text).group()
                    else:
                        if (docu_body_load_date_re.search(total_text) != None):
                            body_text = docu_body_load_date_re.search(total_text).group()
                        else:
                            body_text = ''
                if (docu_date.search(total_text) != None):
                    date_text = docu_date.search(total_text).group()
                else:
                    date_text = ''
            else:
                print(i, "에러")
            res = pd.DataFrame([[publisher, title_publisher_list[i], date_text, body_text]],
                               columns=['Publisher', 'Title', 'Date', 'Body'])
            Master = Master.append(res, ignore_index=True)

        # 전처리
        Master.dropna(subset=['Date'], inplace=True)  # Date 빈칸 제거
        Master['Date'] = pd.to_datetime(Master['Date'])  # 시간 변수로 변환
        Master.drop_duplicates(subset=['Title', 'Body'], inplace=True)  # Title & Body 중복 제거
        Master.drop_duplicates(subset=['Title'], inplace=True)  # Title 중복 제거
        Master.drop_duplicates(subset=['Body'], inplace=True)  # Body 중복 제거
        Master.dropna(subset=['Body'], inplace=True)  # Body 빈칸 제거

        Master.sort_values('Date', inplace=True)
        Master = Master[(Master.Date.dt.year >= 2000) & (Master.Date.dt.year <= 2019)]  # 2000~2019년 만 활용
        if (len(Master) == 0 and curr_number == len(Master)):
            publisher = None
            continue

        Master.reset_index(drop=True, inplace=True)  # Index 재조정
        # Title + Body
        Master['Contents'] = Master.apply(lambda x: x['Title'] + " \n" + x['Body'], axis=1)  # 제목 + 본문
        print(publisher)
        country = mwp_df.loc[mwp_df[mwp_df.columns[0]] == publisher].values[0][1]

        res_ = pd.DataFrame([country], columns=['country'])
        # Master = Master.append(res_, ignore_index=True)
        columnsTitles = ['Publisher', 'Country', 'Title', 'Date', 'Body']
        Master = Master.reindex(columns=columnsTitles)

        if(len(Master) == 0):
            publisher = None
            continue

        Master.to_excel(r"/Users/chanhee.kang/Desktop/sa_paser/test.xlsx", index=None)
        idx += 1