# -*- coding: utf-8 -*-
import os
import pandas as pd


splicing_Depth = 1965

# 定义一个函数，增加重新计算后的厚度列
def get_thickness(x):
    thickness = x['井段End'] - x['井段Start']
    return thickness


PATH = ".\\合并单层表test"
for fileName in os.listdir(PATH):
    if '1单-1' in fileName:
        fileDir1 = PATH + "\\" + fileName
    elif '1单-2' in fileName:
        fileDir2 = PATH + "\\" + fileName

df1 = pd.read_excel(fileDir1, header=2, index='序号')
df1.drop([0], inplace=True)
df1.loc[:, '井 段\n (m)'] = df1['井 段\n (m)'].str.replace(' ', '')  # 消除数据中空格
df1.drop([len(df1)], inplace=True)
df1['井段Start'] = df1['井 段\n (m)'].map(lambda x: x.split("-")[0])
df1['井段End'] = df1['井 段\n (m)'].map(lambda x: x.split("-")[1])
# 表格数据清洗
df1.loc[:, "井段Start"] = df1["井段Start"].str.replace(" ", "").astype('float')
df1.loc[:, "井段End"] = df1["井段End"].str.replace(" ", "").astype('float')

# 截取拼接点以上的数据体
df_temp1 = df1.loc[(df1['井段Start'] <= splicing_Depth), :].copy()#加上copy()可防止直接修改df报错
df_temp1.loc[:, "重计算厚度"] = df_temp1.apply(get_thickness, axis=1)
# print(df_temp1)

#####################################################
df2 = pd.read_excel(fileDir2, header=2, index='序号')
df2.drop([0], inplace=True)
df2.loc[:, '井 段\n (m)'] = df2['井 段\n (m)'].str.replace(' ', '')  # 消除数据中空格
df2.drop([len(df2)], inplace=True)
df2['井段Start'] = df2['井 段\n (m)'].map(lambda x: x.split("-")[0])
df2['井段End'] = df2['井 段\n (m)'].map(lambda x: x.split("-")[1])
# 表格数据清洗
df2.loc[:, "井段Start"] = df2["井段Start"].str.replace(" ", "").astype('float')
df2.loc[:, "井段End"] = df2["井段End"].str.replace(" ", "").astype('float')

#截取拼接点以下的数据体
df_temp2 = df2.loc[(df2['井段Start'] >= splicing_Depth), :].copy()#加上copy()可防止直接修改df报错
df_temp2.reset_index(drop=True, inplace=True)#重新设置列索引
df_temp2.loc[:, "重计算厚度"] = df_temp2.apply(get_thickness, axis=1)
# print(df_temp2)


df_all = df_temp1.append(df_temp2)
df_all.reset_index(drop=True, inplace=True)#重新设置列索引
#对df_all进行操作
df_all.loc[len(df_temp1) - 1, '井 段\n (m)'] = ''.join([str(df_all.loc[len(df_temp1) - 1, '井段Start']), '-', \
                                                  str(df_all.loc[len(df_temp1), '井段Start'])])
print(df_all.loc[len(df_temp1) - 1, '井 段\n (m)'])
df_all.loc[len(df_temp1) - 1, '厚 度\n (m)'] = df_all.loc[len(df_temp1), '井段Start'] - df_all.loc[len(df_temp1) - 1, '井段Start']

# df_all.drop(['井段Start', '井段End', '重计算厚度'], axis=1, inplace=True)
df_all.set_index(["解释\n序号"], inplace=True)
df_all.reset_index(drop=True, inplace=True)#重新设置列索引
#保存为Excel
writer = pd.ExcelWriter('output.xlsx')
df_all.to_excel(writer,'Sheet1')
writer.save()
# print(df_all)