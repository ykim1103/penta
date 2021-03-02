#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 25 2021
@author: LMH @ Penta System Technology

Usage:
    from ExtractTextFromHwp import ExtractText
    ExtractText(target_hwp, destination_text)
    
requirments:
    pyhwp
"""

import os, shutil, re
import datetime as dt
import pandas as pd
import subprocess


def call_subproc(cmdlist):
    print ("subproc cmd list:", cmdlist)
    try:
        p = subprocess.Popen(cmdlist, 
                             stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE,
                             encoding='utf8')
        stdout, stderr = p.communicate()
        
        if stderr:
            raise RuntimeError(stderr)
            
    except Exception as e:
        raise e


def make_quoted_string(string):
    return '"'+string+'"'


def ExtractText(fromfile, tofile=""):
    """
    parameters:
        fromfile : file name with full path for converting hwp to text
        tofile : file name with full path for saving
                 if null, return dataframe 
                 if not null, return file name(tofile)
    """
    type = "pre"

    #################################################################################
    ################## 1. hwp to txt 변환 ###########################################
    #################################################################################   

    cur_path = os.getcwd()
    time_path = "{:%Y%m%d%H%M%S}".format(dt.datetime.now())
    dir_data = os.path.join(cur_path, "temp", time_path)
    os.makedirs(dir_data)
    
    try:
        ##### 1.1. hwp5txt 변환 진행
        rename_doc = os.path.join(dir_data, (type +".hwp"))
        shutil.copyfile(fromfile, rename_doc)
        
        cmdlist = ["hwp5txt", rename_doc, "--output", dir_data + "/pre.txt"]
        call_subproc(cmdlist)

    except Exception as e:
        shutil.rmtree(os.path.join(cur_path, "temp"))
        raise e


    #################################################################################
    ################## 2. 변환된 텍스트 전처리 후 저장 #################################
    #################################################################################      
    flag = ["<표시작>", "<표끝>", "<그림>", ""] # 공백 포함 공백제거 태그
    cell_flag = ["<셀 시작>", "<셀 끝>"] # 셀 태그 

    conv_data = os.path.join(dir_data, (type+".txt"))
    with open(conv_data, 'rt', encoding = "UTF-8") as f:
        pre_text = f.read().splitlines()
        f.close()

    shutil.rmtree(os.path.join(cur_path, "temp"), ignore_errors=True)

    ##### 2.2. 태그 제거 및 전처리 
    pre_data = [x.strip() for x in pre_text if (x.replace(" ", "").strip() not in flag)] 
    number = [no for no, line in enumerate(pre_data) if line.startswith("<셀")] # 셀 태그 위치
    regex_preproc = re.compile('[^ 가-힣|0-9|A-Za-z]+') # 정규식 규칙 정의
    regex_spec = re.compile('^([\(|\[|\{|\<][가-힣|0-9|A-Za-z]{0,3}[\)|\]|\}|\>])|([^가-힣|0-9|A-Za-z.]{0,3}[.])|([^가-힣|0-9|A-Za-z]+)|([가-힣|0-9|A-Za-z]{0,3}[.][\s])|[가-힣|0-9|A-Za-z]{0,3}[\)|\]]')
    regex_bracket = re.compile('[\(|\[|\{|\<|\〈|\《|\「|\｢|\『|\【|\〔|\〈|\❨|\❪|\❬|\❮|\❰|\❲|\❴][\s]{0,1}(별첨|붙임|서식|양식|별표|별지|참조)[가-힣|0-9|A-Za-z|\-\,| ]+[\)|\]|\}|\>|\〉|\》|\」|\｣|\』|\】|\〕|\〗|\〉|\❩|\❫|\❭|\❯|\❱|\❳|\❵]|[0-9|\.|\-]+$')
    
    for j in range(0, len(number), 2): # 셀 태그 위치 (시작, 끝 -> 2개씩)        
        ### 2.2.1 셀 줄넘김 제거시 6바이트 이하이면 줄넘기 제거, 아니면 기존 데이터 유지
        cell_text, k = "", 0 
        for k in range(number[j]+1, number[j+1]):
            cell_text = cell_text + pre_data[k].strip() + " "
        if len(regex_preproc.sub("", str(cell_text).replace(" ", ""))) <=  6 and str(cell_text) != "":
            # 줄넘김 제거 이후 문장 NULL로 남김
            pre_data[number[j]+1] = cell_text.strip() # 셀 첫번째 문장 줄넘김 텍스트로 치환  
            for l in range(number[j]+2, number[j+1]):
                pre_data[l] = ""                   
        ### 2.2.2 셀 내에서 줄넘김 + 특수문자 이면 줄넘김 유지
        i, tmp_cnt = 0, number[j]+1
        for i in range(number[j]+1, number[j+1], tmp_cnt-number[j]): # 셀 시작 ~ 셀 끝
            if re.match(regex_spec, pre_data[i+1].strip()) == None: 
                # 다음문장이 특수문자로 시작하지 않으면
                tmp_cnt = i
                tmp_text = eval("{0}_data".format(type))[i]
                while re.match(regex_spec, pre_data[tmp_cnt+1].strip()) == None:                            
                    # 다음문장이 특수문자로 시작할 때까지
                    tmp_text = tmp_text + " " + pre_data[tmp_cnt+1].strip()
                    pre_data[tmp_cnt+1] = ""
                    tmp_cnt = tmp_cnt + 1
                pre_data[i] = tmp_text

    #################################################################################
    ################## 3. 데이터 전처리 및 규칙 적용 (text, raw, proc, textonly) ##
    #################################################################################                      

    pre_raw = [str(x).strip() for x in pre_data if x not in cell_flag]
    pre_data1 = [regex_bracket.sub("", x).strip() for x in pre_raw if x not in cell_flag] # orgtext
    for (i, text) in enumerate(pre_data1):
        m = re.match(regex_spec, text)
        if m != None:
            pre_data1[i] = pre_data1[i][m.end():].lstrip()
    pre_data2 = [regex_preproc.sub("", str(x).replace(" ", "")) for x in pre_data1 if x not in cell_flag] # textonly
    
                
    #################################################################################
    ################## 4. return용 데이터 생성 #######################################
    ################################################################################# 
    
    data = pd.DataFrame(pre_data1, columns = ["proctext"])
    data["rawtext"] = pre_raw # 원천 텍스트
    data["textonly"] = pre_data2
    data["length"] = data["textonly"].apply(len) # 순수한 텍스트 길이(띄어쓰기, 특수문자 제거)
    
    data = data[data["length"] != 0]    # 텍스트 길이가 0인 데이터 제외
    data = data[["rawtext", "proctext", "textonly"]]
    
    if tofile=="":
        return data
    
    data.to_csv(tofile, index = False, encoding = "UTF-8")
    return tofile