import xlrd
import json
import io
import re
import numpy as np

file_location = "D:/Evision/cleaned.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
masoutis_desc = []
masoutis_cat1 = []
masoutis_cat2 = []
masoutis_cat3 = []
masoutis_cat4 = []
masoutis_brand = []
for i in range(sheet.nrows):
    masoutis_desc.append(sheet.cell_value(i,9))
    masoutis_cat1.append(sheet.cell_value(i,1))
    masoutis_cat2.append(sheet.cell_value(i,3))
    masoutis_cat3.append(sheet.cell_value(i,5))
    masoutis_cat4.append(sheet.cell_value(i,7))
    masoutis_brand.append(sheet.cell_value(i,10))

with io.open('shot35.json',encoding='utf-8') as f:  
    a = json.load(f)
my_words=[]
my_words=[]
for item in a['textAnnotations']:
    item['description'] = re.sub('ά','α',item['description'])
    item['description'] = re.sub('ί','ι',item['description'])
    item['description'] = re.sub('ή','η',item['description'])
    item['description'] = re.sub('ύ','υ',item['description'])
    item['description'] = re.sub('ό','ο',item['description'])
    item['description'] = re.sub('ώ','ω',item['description'])
    item['description'] = re.sub('έ','ε',item['description'])
    my_words.append(item['description'].upper())
    # remove words with less than two letters
del my_words[0]
my_words = [i for i in my_words if len(i) > 3]
my_words = [x for x in my_words if not (x.isdigit() 
                                     or x[0] == '-' and x[1:].isdigit())]

valid_words=[]
sel_k=[]
cnt_found=[]
cnt=0
for i in range(len(my_words)):
    for k in range(len(masoutis_desc)):
        new_data=[]
        for j in range(len(masoutis_desc[k].split(' '))):
            new_data.append(masoutis_desc[k].split(' ')[j])
        if (my_words[i] in new_data)==True:
            cnt=cnt+1
            sel_k.append(k)
            valid_words.append(my_words[i])
    if cnt!=0:
        cnt_found.append(cnt)
    cnt=0

valid_words_unq=[]
cnt = 0
for i in range(len(cnt_found)):
    valid_words_unq.append(valid_words[cnt])
    cnt=cnt+cnt_found[i]

cat2 = [];cat3 = [];cat4= []; desc = []
my_cat2 = [];my_cat3 = []; my_cat4 = []; my_desc=[]
cnt=0
for i in range(len(cnt_found)):
    for j in range(cnt_found[i]):
        cat2.append([])
        cat2[i].append(masoutis_cat2[sel_k[cnt]])
        cat3.append([])
        cat3[i].append(masoutis_cat3[sel_k[cnt]])
        cat4.append([])
        cat4[i].append(masoutis_cat4[sel_k[cnt]])
        desc.append([])
        desc[i].append(masoutis_desc[sel_k[cnt]])
        cnt+=1
    my_cat2.append(np.unique(cat2[i]))
    my_cat3.append(np.unique(cat3[i]))
    my_cat4.append(np.unique(cat4[i]))
    my_desc.append(np.unique(desc[i]))	

my_cat2_1D = [y for x in my_cat2 for y in x]
my_cat2_unq = np.unique(my_cat2_1D)
my_cat3_1D = [y for x in my_cat3 for y in x]
my_cat3_unq = np.unique(my_cat3_1D)
my_cat4_1D = [y for x in my_cat4 for y in x]
my_cat4_unq = np.unique(my_cat4_1D)
my_desc_1D = [y for x in my_desc for y in x]
my_desc_unq = np.unique(my_desc_1D)
count_cat2=[]
for i in range(len(my_cat2_unq)):
    count_cat2.append(my_cat2_1D.count(my_cat2_unq[i]))
count_cat3=[]
for i in range(len(my_cat3_unq)):
    count_cat3.append(my_cat3_1D.count(my_cat3_unq[i]))
count_cat4=[]
for i in range(len(my_cat4_unq)):
    count_cat4.append(my_cat4_1D.count(my_cat4_unq[i]))
count_desc=[]
for i in range(len(my_desc_unq)):
    count_desc.append(my_desc_1D.count(my_desc_unq[i]))
	
count_cat2_sort=sorted(range(len(count_cat2)), key=lambda k: count_cat2[k],reverse=True)
count_cat3_sort=sorted(range(len(count_cat3)), key=lambda k: count_cat3[k],reverse=True)
count_cat4_sort=sorted(range(len(count_cat4)), key=lambda k: count_cat4[k],reverse=True)
count_desc_sort=sorted(range(len(count_desc)), key=lambda k: count_desc[k],reverse=True)
print("Input image is characterized as -", my_cat2_unq[count_cat2_sort[0]],"- in Category2")
print("Input image is characterized as -", my_cat3_unq[count_cat3_sort[0]],"- in Category3")
print("Input image is characterized as -", my_cat4_unq[count_cat4_sort[0]],"- in Category4")
print("Input image is described as -", my_desc_unq[count_desc_sort[0]],"- in Description")