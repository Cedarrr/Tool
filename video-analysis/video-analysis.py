# 参考
# https://blog.csdn.net/baidu_41797613/article/details/121555351

import pandas as pd
import numpy as np


# 1、process video file, get the name of each category in period (output to different files)
type_pd=pd.DataFrame()

video=pd.read_excel("video.xlsx")
market=pd.read_excel("market.xlsx")
START_DATE = 220704
END_DATE = 220710

video = video[video['制作时间'] >= START_DATE]
video = video[video['制作时间'] <= END_DATE]
video_category_list=video['素材类型'].unique()

writer=pd.ExcelWriter('output.xlsx')

for TYPE in video_category_list:

    video_type = video[video['素材类型'] == TYPE]
    video_name_list=video_type['素材名'].unique()

    new = pd.DataFrame()

    for i in video_name_list:
        if (market[market['素材名'].str.contains(i)].empty == False):
            #new=new.append(market[market['素材名']==i])
            new=new.append(market[market['素材名'].str.contains(i)],ignore_index=True)
            new['TYPE']=TYPE
    type_pd=type_pd.append(new)

    new.to_excel(writer,sheet_name=TYPE,index=False)

type_pd.to_excel(writer,sheet_name='ALL',index=False)

# pt=pd.pivot_table(type_pd,index=[u'TYPE',u'素材名'],aggfunc=np.sum,margins=True)
def get_pt(pt):
    new_pt = pd.DataFrame()
    new_pt['消耗'] = pt['消耗']
    new_pt['激活成本'] = round(pt['消耗'] / pt['激活数'], 1)
    new_pt['付费成本'] = round(pt['消耗'] / pt['首日付费人数'], 1)
    new_pt['点击率（%）'] = round(pt['点击数'] / pt['曝光数'], 4) * 100
    new_pt['转化率（%）'] = round(pt['激活数'] / pt['点击数'], 3) * 100
    new_pt['注册率（%）'] = round(pt['注册数'] / pt['激活数'], 3) * 100
    new_pt['首日付费率（%）'] = round(pt['首日付费人数'] / pt['激活数'], 3) * 100
    new_pt['次留率（%）'] = round(pt['次留数'] / pt['激活数'], 3) * 100
    new_pt['首日ROI（%）'] = round(pt['首日付费金额'] / pt['消耗'], 3) * 100
    new_pt['三日ROI（%）'] = round(pt['F3付费金额'] / pt['消耗'], 3) * 100
    new_pt['3/1涨幅'] = round(new_pt['三日ROI（%）'] / new_pt['首日ROI（%）'], 2)
    new_pt['有效率（%）'] = round(pt['有效数'] / pt['激活数'], 3) * 100
    return new_pt

pt_video=pd.pivot_table(type_pd,index=[u'TYPE',u'素材名'],aggfunc=np.sum,margins=True)
#pt_video=pd.pivot_table(type_pd,index=[u'TYPE',u'素材名'],aggfunc=np.sum,margins=True).sort_values(ascending=False,by=['消耗'])
new_pt_video=get_pt(pt_video)
pt_market=pd.pivot_table(market,index=[u'端',u'素材名'],aggfunc=np.sum,margins=True)
new_pt_market=get_pt(pt_market)
# fig=plt.figrue(figsize=(50,50),dpi=1400)
# ax=fig.add_subplot
new_pt_video.to_excel(writer,sheet_name='pt-video',index=True)
new_pt_market.to_excel(writer,sheet_name='pt-market',index=True)
writer.save()

# all_data = pd.pivot_table(new, index=['month'], aggfunc=np.sum, margins=True)
# print(all_data)
print('done')