from define import *
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

query_date = input('请输入查询日期: ')
query_date = datetime.datetime.strptime(query_date,'%Y-%m-%d')
lastyear_date = query_date.replace(query_date.year-1)

#  每月门诊量科室排名
op_num = db.session.query(OutPatient.deptname,OutPatient.NumOP).filter(OutPatient.date==query_date).all()
op_num_lastyear = db.session.query(OutPatient.deptname,OutPatient.NumOP).filter(OutPatient.date==lastyear_date).all()
df1 = pd.DataFrame(op_num,columns=['科室','门诊人次'])
df2 = pd.DataFrame(op_num_lastyear,columns=['科室','同期'])
op_num_df = pd.merge(df1,df2,how='inner',on='科室')
op_num_df.sort_values(by='门诊人次',inplace=True,ascending=False)
op_num_df.plot.bar(x='科室',y=['门诊人次','同期'])
plt.title('本月与同期各科门诊量排名',size=18)
plt.tight_layout()
ax = plt.gca()
ax.set_xticklabels(op_num_df['科室'],rotation=45,ha='right')
plt.show()
# print(op_num)
# print(op_num_lastyear)
# print(query_date,lastyear_date)
# print(op_num_df)

# 全院门诊量
op_sums = []
op_sums_lastyear = []
months = []
for m in range(1,13):
    # print(m)
    op_sum_num = db.session.query(func.sum(OutPatient.NumOP)).filter(OutPatient.date==query_date.replace(month=m)).all()
    op_sum_num_lastyear = db.session.query(func.sum(OutPatient.NumOP)).filter(OutPatient.date==lastyear_date.replace(month=m)).all()
    if op_sum_num[0][0] is None:
        op_sum_num = 0
    else:
        op_sum_num = op_sum_num[0][0]
    op_sum_num_lastyear = op_sum_num_lastyear[0][0]
    months.append(m)
    # print(months)
    op_sums.append(int(op_sum_num))
    op_sums_lastyear.append(int(op_sum_num_lastyear))
# print(op_sums,op_sums_lastyear)
df = pd.DataFrame({'月份':months,'门诊人次':op_sums,'同期':op_sums_lastyear})
plt.subplot(1,1,1)
plt.bar(x=df['月份'],height=df['门诊人次'],label='门诊人次')
plt.plot(df['月份'],df['同期'],marker='o',color='orange',label='同期值')
plt.tight_layout()
plt.xticks(ticks=range(1,13),labels=range(1,13))
plt.xlabel('月份')
plt.ylabel('门诊量')
plt.legend()
plt.title('全院门诊病人量',size=18)
plt.show()

# 门诊病历书写率\门诊病历甲级率
op_num = db.session.query(OutPatient.deptname,OutPatient.NumOP).filter(OutPatient.date==query_date).all()
op_rec_num = db.session.query(OutPatient.deptname,OutPatient.NumOPRec).filter(OutPatient.date==query_date).all()
op_checkrec_num = db.session.query(OutPatient.deptname,OutPatient.CheckNumOPRec).filter(OutPatient.date==query_date).all()
op_arec_num = db.session.query(OutPatient.deptname,OutPatient.ANumOPRec).filter(OutPatient.date==query_date).all()
df1 = pd.DataFrame(op_num,columns=['科室','门诊人次'])
df2 = pd.DataFrame(op_rec_num,columns=['科室','病历数'])
df3 = pd.DataFrame(op_rec_num,columns=['科室','检查病历数'])
df4 = pd.DataFrame(op_rec_num,columns=['科室','甲级病历数'])
op_rec_df = pd.merge(df1,df2,how='inner',on='科室')
op_rec_df = pd.merge(op_rec_df,df3,how='inner',on='科室')
op_rec_df = pd.merge(op_rec_df,df4,how='inner',on='科室')
op_rec_df['门诊病历书写率'] = op_rec_df['病历数']/op_rec_df['门诊人次']
op_rec_df['门诊病历甲级率'] = op_rec_df['甲级病历数']/op_rec_df['检查病历数']
op_rec_df.sort_values(by='门诊病历书写率',inplace=True,ascending=False)
op_rec_df.plot.bar(x='科室',y=['门诊病历书写率','门诊病历甲级率'])
plt.title('本月门诊病历书写率/甲级率科室排名',size=18)
plt.tight_layout()
ax = plt.gca()
ax.set_xticklabels(op_rec_df['科室'],rotation=45,ha='right')
plt.show()
print(op_rec_df)

# 住院病历甲级率/按时归档率
ip_num = db.session.query(InPatient.deptname,InPatient.NumIP).filter(InPatient.date==query_date).all()
ip_oanum = db.session.query(InPatient.deptname,InPatient.OANumIPRec).filter(InPatient.date==query_date).all()
ip_checkrec_num = db.session.query(InPatient.deptname,InPatient.CheckNumIPRec).filter(InPatient.date==query_date).all()
ip_arec_num = db.session.query(InPatient.deptname,InPatient.ANumIPRec).filter(InPatient.date==query_date).all()
df1 = pd.DataFrame(ip_num,columns=['科室','出院人次'])
df2 = pd.DataFrame(ip_oanum,columns=['科室','迟交病历数'])
df3 = pd.DataFrame(ip_checkrec_num,columns=['科室','检查病历数'])
df4 = pd.DataFrame(ip_arec_num,columns=['科室','甲级病历数'])
ip_rec_df = pd.merge(df1,df2,how='inner',on='科室')
ip_rec_df = pd.merge(ip_rec_df,df3,how='inner',on='科室')
ip_rec_df = pd.merge(ip_rec_df,df4,how='inner',on='科室')
ip_rec_df['住院病历甲级率'] = ip_rec_df['甲级病历数']/ip_rec_df['检查病历数']
ip_rec_df['住院病历按时归档率'] = (ip_rec_df['出院人次']-ip_rec_df['迟交病历数'])/ip_rec_df['出院人次']
ip_rec_df.sort_values(by='住院病历按时归档率',inplace=True,ascending=False)
ip_rec_df.plot.bar(x='科室',y=['住院病历按时归档率','住院病历甲级率'])
plt.title('本月住院病历按时归档率/甲级率科室排名',size=18)
plt.tight_layout()
ax = plt.gca()
ax.set_xticklabels(ip_rec_df['科室'],rotation=45,ha='right')
plt.show()
print(ip_rec_df )

# 临床路径入径率
ip_num = db.session.query(InPatient.deptname,InPatient.NumIP).filter(InPatient.date==query_date).all()
path_num = db.session.query(InPatient.deptname,InPatient.NumPath).filter(InPatient.date==query_date).all()

df1 = pd.DataFrame(ip_num,columns=['科室','出院人次'])
df2 = pd.DataFrame(path_num,columns=['科室','入径人数'])

path_df = pd.merge(df1,df2,how='inner',on='科室')

path_df['临床路径入径率'] = path_df['入径人数']/path_df['出院人次']

path_df.sort_values(by='临床路径入径率',inplace=True,ascending=False)
path_df.plot.bar(x='科室',y='临床路径入径率')
plt.title('本月临床路径入径率科室排名',size=18)
plt.tight_layout()
ax = plt.gca()
ax.set_xticklabels(path_df['科室'],rotation=45,ha='right')
plt.show()
print(path_df )