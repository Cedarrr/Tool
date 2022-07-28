import pandas as pd
data_loc='80.xlsx'
data = pd.read_excel(data_loc)
out=pd.DataFrame()
OUTPUT_NAME="OUTPUT+"+data_loc

out['账户名称']=data['账户名称']
out['日期']=data['日期']
out['账户消耗']=data['账户总支出(¥)']
out['实际消耗']=round(out['账户消耗']/1.06,2)
out['账户余额']=data['总余额(¥)']
out['加款转账情况']=data['账户总转入(¥)']-data['账户总转出(¥)']
out['加款转账情况']=out['加款转账情况'].replace(0,"")
out.sort_values('日期',inplace=True)
out_clear=out.drop(out[(out['账户消耗']==0)&(out['加款转账情况']=="")].index)

writer = pd.ExcelWriter(OUTPUT_NAME)
out_clear.to_excel(writer,"out_clear",index=False)
out.to_excel(writer,"out",index=False)
data.to_excel(writer,"data",index=False)
writer.save()

print("done")