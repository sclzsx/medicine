from cmath import nan
import numpy as np
import pandas as pd
import docx
import sys
import re

rank = 500

data = np.array(pd.read_excel('各症状及编码.xlsx', sheet_name='各症状', header=None))
label = np.array(pd.read_excel('各症状及编码.xlsx', sheet_name='各症状用药替换数字'))

label = label[:, 3:6]

medicine_single = dict()
medicine_soup = dict()
for a in range(label.shape[0]):
    t = label[:, 2][a]
    if isinstance(t, str):
        # print(t)
        medicine_soup.setdefault(label[:, 1][a], label[:, 0][a])
    else:
        medicine_single.setdefault(label[:, 1][a], label[:, 0][a])

# print(medicine_single)
# print()
# print(medicine_soup)

# for k,v in medicine_soup.items():
#     print(k,v)

name = data[0, :]
data = data[1:, :]
# print(data.shape, label.shape)

data_single = data.copy()

data_soup = data.copy()
# print(data2.shape)

h, w = data.shape
for i in range(h):
    for j in range(w):
        data_single[i, j] = nan
        data_soup[i, j] = nan

        info = data[i, j]
        # print(info)

        if not isinstance(info, str):
            continue

        info = info.replace(',', ' ').replace('，', ' ').replace(':', ' ').replace('：', ' ').replace(';', ' ').replace(
            '；', ' ').replace('。', ' ').replace('(', ' ').replace(')', ' ')
        info = re.sub('[\u4e00-\u9fa5]', ' ', info)
        info = re.sub('[a-zA-Z]', ' ', info)
        info = info.strip()
        info = info.split(' ')

        nums_single = []
        nums_soup = []
        for num in info:
            if len(num) >= 1:
                if len(num) <= 3:

                    num = int(num)

                    if num in medicine_single and num not in medicine_soup:
                        nums_single.append(num)

                    if num not in medicine_single and num in medicine_soup:
                        nums_soup.append(num)
                else:

                    if len(num) == 6:
                        num1 = int(num[:3])
                        num2 = (num[3:])
                        
                        if num1 in medicine_single and num1 not in medicine_soup:
                            nums_single.append(num1)
                        
                        if num1 not in medicine_single and num1 in medicine_soup:
                            nums_soup.append(num1)

                        if num2 in medicine_single and num2 not in medicine_soup:
                            nums_single.append(num2)

                        if num2 in medicine_single and num2 not in medicine_soup:
                            nums_soup.append(num2)

                    else:
                        print('unexpected num:', num)


        # print(len(nums_single), len(nums_soup))
        if len(nums_single) > 0:
            
            data_single[i, j] = nums_single
        
        if len(nums_soup) > 0:
            data_soup[i, j] = nums_soup

        
for i in range(h):
    for j in range(w):
        a = data_single[i, j]
        b = data_soup[i, j]
        if isinstance(a, list):
            for aa in a:
                assert aa in medicine_single and aa not in medicine_soup
        if isinstance(b, list):
            for bb in b:
                assert bb in medicine_soup and bb not in medicine_single

# print(data_single)

with open('single_all_symptom_' + 'rank' + str(rank) + '.txt', 'w') as f:

    counter = medicine_single.copy()
    for k in counter:
        counter[k] = 0
    used = set()

    for j in range(w):

        info = data_single[:, j]
        # print(info)
        for obj in info:
            if isinstance(obj, list):

                # print(obj)

                for num in obj:
                    
                    # print(num)
                    # if num in medicine_soup:
                    #     print('sss', num)
                        # pass
                    # assert num not in medicine_soup

                    used.add(num)
                    # print(num)
                    if num in counter:
                        # print(counter[num])
                        counter[num] = counter[num] + 1
                        # print(counter[num])

    counter_sorted = sorted(counter.items(), key=lambda x: x[1], reverse=True)

        # print(type(counter_sorted))

        # print(used)

    used_real = []
    for t in used:
        used_real.append(medicine_single[t])

    # print(used_real)

    cnt = 0

    

    print('所有症状：', [name[k] for k in range(len(name))], file=f)
    print('涉及单药：', used_real, file=f)
    print('涉及单药门数：', len(used), file=f)
    print('排名前{}的单药有：'.format(rank), file=f)
    for k, v in counter_sorted:
        # print(medicine[k])
        if v > 0:
            if cnt > rank:
                break
            print('单药:{}, 频次:{}, 频率(%):{}'.format(medicine_single[k], v, np.round(v * 100 / len(used), 2)), file=f)
            cnt = cnt + 1
    print('\n', file=f)



with open('soup_all_symptom_' + 'rank' + str(rank) + '.txt', 'w') as f:

    counter = medicine_soup.copy()
    for k in counter:
        counter[k] = 0
    
    # print(counter)

    used = set()

    for j in range(w):


        info = data_soup[:, j]
        # print(info)
        for obj in info:
            if isinstance(obj, list):

                # print(obj)

                for num in obj:
                    
                    # print(num)
                    # if num in medicine_soup:
                    #     print('sss', num)
                        # pass
                    # assert num not in medicine_soup

                    used.add(num)
                    # print(num)
                    if num in counter:
                        # print(counter[num])
                        counter[num] = counter[num] + 1
                        # print(counter[num])

    counter_sorted = sorted(counter.items(), key=lambda x: x[1], reverse=True)

        # print(type(counter_sorted))

        # print(used)

    used_real = []
    for t in used:
        used_real.append(medicine_soup[t])

        # print(used_real)

    cnt = 0

    print('所有症状：', [name[k] for k in range(len(name))], file=f)
    print('涉及方剂：', used_real, file=f)
    print('涉及方剂门数：', len(used), file=f)
    print('排名前{}的方剂有：'.format(rank), file=f)
    for k, v in counter_sorted:
        # print(medicine[k])
        if v > 0:
            if cnt > rank:
                break
            print('方剂:{}, 频次:{}, 频率(%):{}'.format(medicine_soup[k], v, np.round(v * 100 / len(used), 2)), file=f)
            cnt = cnt + 1
    print('\n', file=f)