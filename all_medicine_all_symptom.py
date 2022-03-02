import numpy as np
import pandas as pd
import docx
import sys
import re

rank = 500

data = np.array(pd.read_excel('各症状及编码.xlsx', sheet_name='各症状', header=None))
label = np.array(pd.read_excel('各症状及编码.xlsx', sheet_name='各症状用药替换数字'))

label = label[:, 3:6]

medicine = dict()
for a in range(label.shape[0]):
    medicine.setdefault(label[:, 1][a], label[:, 0][a])

# print(medicine)

name = data[0, :]
data = data[1:, :]
# print(data.shape, label.shape)

data2 = data[:]
# print(data2.shape)

h, w = data.shape
for i in range(h):
    for j in range(w):
        info = data[i, j]

        if not isinstance(info, str):
            continue

        info = info.replace(',', ' ').replace('，', ' ').replace(':', ' ').replace('：', ' ').replace(';', ' ').replace(
            '；', ' ').replace('。', ' ').replace('(', ' ').replace(')', ' ')
        info = re.sub('[\u4e00-\u9fa5]', ' ', info)
        info = re.sub('[a-zA-Z]', ' ', info)
        info = info.strip()
        info = info.split(' ')

        nums = []
        for num in info:
            if len(num) >= 1:
                if len(num) <= 3:
                    num = int(num)
                    nums.append(num)
                else:
                    if len(num) == 6:
                        num1 = int(num[:3])
                        num2 = (num[3:])
                        # print('change unexpected num into:', num1, num2)
                    else:
                        print('unexpected num:', num)


        data2[i, j] = nums

with open('all_medicine_all_symptom_ ' + 'rank' + str(rank) + '.txt', 'w') as f:

    counter = dict()
    for l in label[:, 1]:
        counter.setdefault(l, 0)

    used = set()

    for j in range(w):

        info = data2[:, j]
        for obj in info:
            if not np.isnan(obj).any() and isinstance(obj, list):

                # print(obj)
                for num in obj:

                    used.add(num)
                    # print(num)
                    if num in counter:
                        # print(counter[num])
                        counter[num] = counter[num] + 1
                        # print(counter[num])

    counter_sorted = sorted(counter.items(), key=lambda x: x[1], reverse=True)

    # print(counter_sorted)

    used_real = []
    for t in used:
        used_real.append(medicine[t])

    cnt = 0

    

    print('所有症状：', [name[k] for k in range(len(name))], file=f)
    print('所有涉及药物：', used_real, file=f)
    print('涉及药物门数：', len(used), file=f)
    print('排名前{}的药物有：'.format(rank), file=f)
    for k, v in counter_sorted:
        # print(medicine[k])
        if v > 0:
            if cnt > rank:
                break
            print('药物:{}, 频次:{}, 频率(%):{}'.format(medicine[k], v, np.round(v * 100 / len(used), 2)), file=f)
            cnt = cnt + 1
    print('\n', file=f)
