# -*- coding: utf-8 -*-
"""
Created on Fri Sep  4 07:30:33 2020

@author: EastmanE
"""

import os
import datetime as dt
import pandas as pd
import openpyxl as xl
import numpy as np
import xml.etree.ElementTree as ET
from MRI_settings import filepath, paramlist, newparamlist, folderpath
import win32com.client as win32
import xlsxwriter 

year = dt.datetime.today().strftime('%Y')


# =============================================================================
# Read files (must be xlsx file)
# =============================================================================

tree = ET.parse(filepath)                                         #Parse the xml file
root = tree.getroot()                                                   

taglist = [elem.tag for elem in root.iter()]                        #Write the keywords to a list
textlist = [elem.text for elem in root.iter()]                      #Write the values to a list
attrlist = [elem.attrib for elem in root.iter()]

df = pd.DataFrame()                                                 #Write the lists to a df for
df[0] = taglist                                                     #easier analysis
df[1] = textlist
df[2] = attrlist

machinename = df.loc[df[0] == 'HeaderTitle', 1].tolist()[0]


# #Find 
# tab_sections = df.loc[df[0] == 'region'].index.tolist()
# tab_sections.append(df.loc[df[0] == 'PrintProtocol'].index.tolist()[0])

# exam_list = []
# type_list = []
# program_list = []
# sequence_list = []

# for idx1, section1 in enumerate(tab_sections[:-1]):
#     startrow1 = section1
#     stoprow1 = tab_sections[idx1+1]
#     df1 = df.loc[startrow1:stoprow1, :]
#     tabname = df1.loc[startrow1,2]['name']
#     print(tabname)

#     exam_sections = df1.loc[df[0] == 'NormalExam_dot_engine'].index.tolist()
#     seq_sep = df.loc[df[0] == 'PrintProtocol'].index.tolist()
    
#     #Start of sequence details is end of seq summary
#     last_row = seq_sep[0]
#     exam_sections.append(stoprow1)
    
#     print(exam_sections)
#     #Sep Soft Tissue from MRA from Face
#     for idx, section in enumerate(exam_sections[:-1]):
#         examname = df.loc[section, 2]['name']
#         startrow = section
#         stoprow = exam_sections[idx+1]
#         df2 = df.loc[startrow:stoprow, :]
#         #Sep RTN from optional from test
#         program_sections = df2.loc[df2[0] == 'program'].index.tolist()
#         program_sections.append(stoprow)
        
#         for idx2, section2 in enumerate(program_sections[:-1]):
#             programname = df.loc[section2, 2]['name']
#             startrow2 = section2+1
#             stoprow2 = program_sections[idx2+1]
#             df3 = df2.loc[startrow2:stoprow2, :]
#             counter = 0
#             for rowidx, row in enumerate(df3.index[:-1]):
#                 if 'localizer' in str(df.loc[row, 2]['name']).lower():
#                     continue
#                 else:
#                     seqname = str(df.loc[row, 2]['name']) + ' (' + str(counter) + ')'
#                     counter +=1
#                     exam_list.append(tabname)
#                     type_list.append(examname)
#                     program_list.append(programname)
#                     sequence_list.append(seqname)

# collectdf = pd.DataFrame(data = {'Exam': exam_list, 'Type': type_list, 'Program':program_list, 'Sequence Name':sequence_list})

# print('collectdf complete')

# collectdf.to_excel(finalpath, sheet_name = tabname, index = False)      


headerprotpath = []

headerprotpath = df.loc[(df[0] == 'HeaderProtPath')].index.tolist()

# for m in dfcheck.index:
#     if (not 'localizer' in dfcheck.loc[m,1] )& (not 'Localizer' in dfcheck.loc[m,1] ):
#         headerprotpath.append(m)
# headerprotpath = [i for i in headerprotpath if 'localizer' not in i if'Localizer' not in i]
headerprotpath.append(len(df))   
lv_collectdf = pd.DataFrame()

print('starting reformat')



exam_list = []
type_list = []
program_list = []
sequence_list = []
cat_list = []
label_list = []
value_list = []
# keeptrack = ''

for idx, header in enumerate(headerprotpath[:-1]):
    protpath = df.loc[header, 1]
    pathsplit = protpath.split('\\')
    seqname = pathsplit[-1]
    program = pathsplit[-2]
    seqtype = pathsplit[-3]
    examtype = pathsplit[-4]
    if 'Substep:' in df.loc[header+1, 1]:
        substep = df.loc[header+1, 1].split(':')[-1].strip()
        kernel = df.loc[header+1, 1].split('|')[0].split(':')[-1].strip()
        seqname = seqname + ' (' + substep + ')' + ' ('+kernel +')'
    else:
        kernel = df.loc[header+1, 1].split(':')[-1].strip()
        seqname = seqname + ' (' + kernel + ')'
    
    if 'localizer' in seqname.lower():
        continue
    else:
        pass
    
    tavalue = df.loc[header+1, 1].split(':', 1)[1].strip().split(' ')[0]
    exam_list.append(examtype)
    type_list.append(seqtype)
    program_list.append(program)
    sequence_list.append(seqname)
    cat_list.append('NA')
    label_list.append('TA')
    value_list.append(tavalue)

    
    # if seqname != keeptrack:
    #     keeptrack = seqname
    #     keeptrackcount = 0
    #     seqname = seqname + ' (' + str(keeptrackcount) +')'
    # else:
    #     keeptrackcount +=1
    #     seqname = seqname + ' (' + str(keeptrackcount) +')'
    # print(seqname)

    
    startidx = header
    stopidx = headerprotpath[idx+1]
    seqdf = df.loc[startidx:stopidx, :]
    
    #Break down into different card categories
    cardidx = seqdf.loc[seqdf[0] == 'Card'].index.tolist()
    cardidx.append(seqdf.index[-1]+1)
    
    for idx2, card in enumerate(cardidx[:-1]):

        cardname = seqdf.loc[card, 2]['name']
        startidx2 = card
        stopidx2 = cardidx[idx2+1]
        carddf = seqdf.loc[startidx2:stopidx2, :]
        carddf = carddf.loc[carddf[1] != '\n']
        
        labels = carddf.loc[carddf[0] == 'Label'][1].tolist()
        valueandunit = carddf.loc[carddf[0] == 'ValueAndUnit'][1].tolist()
    
        for l, v in zip(labels, valueandunit):
            exam_list.append(examtype)
            type_list.append(seqtype)
            program_list.append(program)
            sequence_list.append(seqname)
            cat_list.append(cardname)
            label_list.append(l)
            value_list.append(v)

lv_collectdf = pd.DataFrame(data = {'Exam':exam_list, 'Type':type_list, 'Program':program_list, 'Sequence Name':sequence_list, 'Category':cat_list, 'Label':label_list, 'Value':value_list})
print('reformat complete')

lv_collectdf_noloc = lv_collectdf.loc[~lv_collectdf['Sequence Name'].str.contains('localizer|Localizer')]

lv_collectdf_noloc = lv_collectdf_noloc[lv_collectdf_noloc['Label'].isin(paramlist)]
    
lv_collectdf_noloc.drop('Category', axis = 1, inplace = True)        
lv_collectdf_noloc.drop_duplicates(inplace = True)

lv_collectdf_noloc['NewName'] = lv_collectdf_noloc['Exam'] + ' ' + lv_collectdf_noloc['Type']+ ' ' + lv_collectdf_noloc['Program'] + ' ' + lv_collectdf_noloc['Sequence Name']


excel = win32.Dispatch('Excel.Application')

for regionname in list(set(lv_collectdf_noloc['Exam'])):
    new_lv_collectdf_noloc = lv_collectdf_noloc.loc[lv_collectdf_noloc['Exam'] == regionname]
    if '/' in regionname:
        newregionname= regionname.replace('/', '_')
        regionfilename = newregionname + '.xlsx'
    else:
        regionfilename = regionname + '.xlsx'
    
    seqlist = []
    for i in new_lv_collectdf_noloc['NewName'].tolist():
        if i not in seqlist:
            seqlist.append(i)
        else:
            pass
    
    reformatdf = pd.DataFrame()
    for seq in seqlist:
        appendlist = []
        seqdf=new_lv_collectdf_noloc.loc[new_lv_collectdf_noloc['NewName']==seq]
        exam = list(set(seqdf['Exam']))[0]
        typeval = list(set(seqdf['Type']))[0]
        program = list(set(seqdf['Program']))[0]
    
        for param in paramlist:
            if param not in seqdf['Label'].tolist():
                appendlist.append('')
            else:
                appendlist.append(seqdf.loc[seqdf['Label'] == param]['Value'].tolist()[0])     
        appenddf = pd.DataFrame(data = [appendlist], columns = paramlist)
        seq = seq.replace(program, '', 1)
        seq = seq.replace(exam, '', 1)
        seq = seq.replace(typeval, '', 1)
        seq = seq.strip()
        
        appenddf.insert(0, 'Sequence', seq)
    
        appenddf.insert(0, 'Program', program)
        appenddf.insert(0, 'Type', typeval)
        appenddf.insert(0, 'Exam', exam)
        
        reformatdf = pd.concat([reformatdf, appenddf], ignore_index = False)
    
    
    currentfile = os.path.join(folderpath, regionfilename)
    with pd.ExcelWriter(currentfile, engine = 'xlsxwriter') as writer:
        reformatdf['FullName'] = reformatdf['Type'] + '- ' + reformatdf['Program']
        
        programs = list(dict.fromkeys(reformatdf['FullName']))
        checkprog = []
        workbook = writer.book
        
        proglist = []
        for prog in programs:
            programdf = reformatdf.loc[reformatdf['FullName'] == prog]
            exam_name = programdf['Program'].tolist()[0]
            programdf.drop(['Exam', 'Type', 'FullName', 'Program'], axis=1, inplace = True)
            prog = prog.replace('/', ' ')
            prog = prog.replace('?', '')

            if prog.lower() not in checkprog:
                checkprog.append(prog.lower())
            else:
                prog = '(Dup) ' + prog
            
            for f1, f2 in zip(paramlist, newparamlist):       
                programdf = programdf.rename(columns = {f1:f2})
            
            
            programdf['# of slices'] = programdf['Slices'] + programdf['# of slices']
            programdf.drop('Slices', axis = 1, inplace = True)
            
            progtab = prog[:31]
            if progtab not in proglist:
                proglist.append(progtab)
            else:
                proglist.append(progtab)
                prognum = proglist.count(progtab)
                prognum = '(' + str(prognum) + ')'
                progtab = progtab[:-4]
                progtab = progtab + prognum
                
            programdf.to_excel(writer, sheet_name = progtab, index = False) 
            worksheet = writer.sheets[progtab]
            border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
            worksheet.write(0, 15, prog)
            worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(programdf), len(programdf.columns)-1), {'type': 'no_errors', 'format': border_fmt})
            
        writer.save()




    wb = excel.Workbooks.Open(Filename = currentfile)
    for ll in wb.Worksheets:
        ll.Columns('A:Z').AutoFit()    # wb.Save()
    wb.Save()
    wb.Close()
    print(regionname)
excel.Quit()        


# reformatdf.to_excel(nbpath, sheet_name = tabname, index = False)        


