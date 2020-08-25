#!/usr/bin/env python
# coding: utf-8

import os
import copy
import pandas

from os.path import join
from pandas.core.frame import DataFrame
from MyPythonDocx import *


def cal_va(df):
    
#     df = DataFrame(page[1:], columns=page[0])
    severity = ['嚴重', '高', '中', '低', '無']
    vas = []

    for idx in range(5):
        mask = df['嚴重程度'] == severity[idx]
        tmp = df[mask][['弱點名稱', '弱點描述']].values.tolist()
        vas.append([])
        for name in tmp:
            if name and name not in vas[idx]:
                vas[idx].append(name)
    
#     print(vas)
    return vas            


def cal_risk_cnt(page):
    
    try:
        df = DataFrame(page[1:], columns=page[0])
    except:
        df = page
    
    severity = ['嚴重', '高', '中', '低', '無']
    cnts = []    
    
    for idx in range(5):
        mask = df['嚴重程度'] == severity[idx]
        cnts.append(df[mask].shape[0])
        
#         total = set(tuple(x) for x in df[mask]['弱點名稱'])
#         print(len(total))

#     print(cnts)
    return cnts


def check_va(data, data2):
    
    # IP位址, Service Port, 弱點ID
    total = ((row[0], row[3], row[4]) for row in data)
    total2 = ((row[0], row[3], row[4]) for row in data2)

    set1, set2 = set(total), set(total2)

    not_repair_va = set1 & set2

    new_va = set2 - not_repair_va

    repaired_va = set1 - not_repair_va

    elem = ('IP位址', 'Service Port', '弱點ID')
    new_va.add(elem)
    repaired_va.add(elem)
    
    return list(new_va), list(not_repair_va), list(repaired_va)


def va_to_row(va, data):

    res = []
    tmp_va = copy.deepcopy(va)
    
    for row in data:
        for cod in tmp_va:
            if row[0] == cod[0] and row[3] == cod[1] and row[4] == cod[2]:
                res.append(row)
                tmp_va.remove(cod)
                break
                
    return res


def doc_va_table(doc, data, idx):
    
    title = ['風險等級', '風險名稱', '風險簡述']
    severity = ['嚴重', '高', '中', '低', '無']
    vas = cal_va(data)

    tb = doc.tables[idx]
    for i, row in enumerate(tb.rows):
        if not i:
            continue
        Table.remove_row(tb, row)
    
    for i in range(6):
        if not i:
            for j in range(3):
                cell = tb.cell(i, j)
                cell.text = title[j]
                Table.set_cell_color(cell, PURPLE)
                Parag.set_font(cell.paragraphs, size=Pt(12), name=u'標楷體')
        else:
            for x in vas[i-1]:
                row_cells = tb.add_row().cells
                row_cells[0].text = severity[i-1]
                Parag.set_font(row_cells[0].paragraphs, size=Pt(12), name=u'標楷體')
                row_cells[1].text = x[0]
                row_cells[2].text = x[1]
                for p in row_cells[2].paragraphs:
                    p.paragraph_format.alignment = 0
#                 print(x[0], ' finished!')
                
    Table.set_content_font(tb, size=Pt(12), name=u'標楷體')
    
    Table.col_widths(tb, 1.5, 2.7, 5)


def doc_risk_cnt_copare(doc, data, data2):
    
    cnts = cal_risk_cnt(data)
    cnts2 = cal_risk_cnt(data2)
#     cnts3 = cal_risk_cnt(data_decrease)
    
    tbl = doc.tables[2]
    sums = [0, 0, 0]
    for i in range(5):
        tbl.cell(i+1, 1).text = str(cnts[i])
        tbl.cell(i+1, 2).text = str(cnts2[i])
        
        sub = cnts[i] - cnts2[i]
        tbl.cell(i+1, 3).text = str(sub)
        
#         res = cnts[i]-cnts2[i] if cnts[i]>cnts2[i] else 0
#         tbl.cell(i+1, 3).text = str(res)
        
        sums[0] = sums[0] + cnts[i]
        sums[1] = sums[1] + cnts2[i]
        sums[2] = sums[2] + sub
    
    for i in range(3):
        tbl.cell(6, i+1).text = str(sums[i])
        
    Table.set_content_font(tbl, size=Pt(12), name=u'標楷體')
    
    for p in doc.paragraphs:
        if 'decrease_risk_cnt' in p.text:
            p.text = p.text.replace('decrease_risk_cnt', str(sums[2]))


def ExcelToWord_second(word, excel, excel2, sheet, sheet2, date, consultant):
    
    ip = word.split('/')[-1].replace('.docx','')
    print(ip)
    
    month, day = date.split('/')
    doc = Document(word)
    
    df_excel = pandas.read_excel(excel, sheet_name=None)
    data = df_excel[sheet]
    
    mask = data['IP位址'] == ip
    new_data = data[mask]
    
    if not new_data.empty:

        df_excel2 = pandas.read_excel(excel2, sheet_name=None)
        data2 = df_excel2[sheet2]
        
        li_data = data.values.tolist()
        li_data.insert(0, data.columns)
        
        li_data2 = data2.values.tolist()
        li_data2.insert(0, data2.columns)

        _, not_repair_va, _ = check_va(li_data, li_data2)
        li_new_data2 = va_to_row(not_repair_va, li_data)
        df_new_data2 = DataFrame(li_new_data2[1:], columns=li_new_data2[0])
        new_data2 = df_new_data2[df_new_data2['IP位址'] == ip]

        
        doc_risk_cnt_copare(doc, new_data, new_data2)    
        
        if new_data2.empty:
            new_data2 = DataFrame(columns=('嚴重程度', '弱點名稱', '弱點描述'))
            for i, severity in enumerate(['高', '中', '低']):
                new_data2.loc[i] = [severity, '無', '-']
        
        doc_va_table(doc, new_data2, 1)
            

        for p in doc.paragraphs:
            row = p.text
            
            if 'MMM' in row:
                row = row.replace('MMM', month)
                p.text = row

            if 'DDD' in row:
                row = row.replace('DDD', day)
                p.text = row

            if 'second_technical_consultant' in row:
                p.text = row.replace('second_technical_consultant', consultant)

            if 'second_detect_ip' in row:
                p.text = row.replace('second_detect_ip', ip)                    

            if 'refer_excel_2' in row:
                p.text = row.replace('refer_excel_2', excel2.split('/')[-1])

        doc.save(word)


def ExcelToWord_first(word, excel, sheet, save, date, consultant):
    
    file_name = excel.split('/')[-1].split('.')[0]
    file_path = join(save, file_name)
    try:
        os.mkdir(file_path)
    except:
        pass
    month, day = date.split('/') 
    doc = Document(word)

    df_excel = pandas.read_excel(excel, sheet_name=None)
    data = df_excel[sheet]
    li_data = data.values.tolist()
    li_data.insert(0, data.columns)
    ip_list = df_excel['主機']['IP位址']

    for ip in ip_list:
        mask = data['IP位址'] == ip
        new_data = data[mask]
        
        if not new_data.empty:
            doc = Document(word)
            doc_va_table(doc, new_data, 0)
                
            for p in doc.paragraphs:
                row = p.text
            
                if 'mmm' in row:
                    row = row.replace('mmm', month)
                    p.text = row

                if 'ddd' in row:
                    row = row.replace('ddd', day)
                    p.text = row

                if 'first_technical_consultant' in row:
                    p.text = row.replace('first_technical_consultant', consultant)

                if 'first_detect_ip' in row:
                    p.text = row.replace('first_detect_ip', ip)


                if 'refer_excel_1' in row: 
                    p.text = row.replace('refer_excel_1', excel.split('/')[-1])

            print(ip)
            doc.save(join(file_path, ip+'.docx'))
    
    os.remove('tmp_template.docx')


def ExcelToWord_doc(word, company, date, tool_name, tool_version):

    month, day = date.split('/')
    doc = Document(word)
    
    first_flag = True
    for p in doc.paragraphs:
        row = p.text

        if 'mmmm' in row:
            row = row.replace('mmmm', month)
            Parag.set_run_font(p, 20)

        if 'dddd' in row:
            row = row.replace('dddd', day)
            p.text = row
            Parag.set_run_font(p, 20)

        if 'tool_name' in row:
            p.text = row.replace('tool_name', tool_name)
            
        if 'tool_version' in row:
            p.text = row.replace('tool_version', tool_version)

        if 'OOOOO' in row:
            p.text = row.replace('OOOOO', company)
            if first_flag:
                Parag.set_run_font(p, 22)
                first_flag = False
            else:
                Parag.set_run_font(p, 20)
    
    doc.save('tmp_template.docx')


if __name__ == '__main__':
    word = '/Volumes/DATA/GoodleDrive/Internship/公司資料/模板/proto_excel_工業局.docx'
    excel = '/Volumes/DATA/GoodleDrive/Internship/公司資料/初掃 xlsx/1216 全聯初掃報告.xlsx'
    excel2 = '/Volumes/DATA/GoodleDrive/Internship/公司資料/複掃 xlsx/0207全聯複掃報告.xlsx'
#     excel = '/Users/yucc/Documents/test/豆腐食品初掃報告 (1).xlsx'
    sheet, sheet2 = 'All', 'All'
    save = '/Users/yucc/Desktop/'
    
    company, cover_date, tool_name, tool_version = '豆腐食品', '05/06', 'Rapid7 NEXPOSE', 'unknown'
    first_date, consultant = '05/07', 'Anthony'
    second_date, consultant2 = '05/10', 'Rick'
    
    ExcelToWord_doc(word, company, cover_date, tool_name, tool_version)
    ExcelToWord_first('tmp_template.docx', excel, sheet, save, first_date, consultant)
    
    print('--------------------------')
    save = '/Users/yucc/Desktop/1216 全聯初掃報告'
    
    for f in os.listdir(save):
        f_path = join(save, f)
        ExcelToWord_second(f_path, excel, excel2, sheet, sheet2, second_date, consultant2)

#     ExcelToWord_final(excel, excel2, sheet, word, save)




