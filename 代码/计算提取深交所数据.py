import xlwings as xw
import pandas as pd
from dbfread import DBF
import datetime
import sys
import os
def write_row(fp,sheet,row,data):
    """
    写入一行
    :param fp:    excel文件
    :param sheet: sheet name 或 sheet index
    :param row:   哪一行， 如 'A1'
    :param data:  list，如['test','hello',...]
    :return:
    """
    wb=xw.Book(fp)
    if isinstance(sheet,str):
        sht=wb.sheets(sheet)
    else:
        sht=wb.sheets[sheet]
    sht.range(row).value=data
    wb.save()


def SJSYE_get_Data(sjsye_dbf_path,excel_path):
    """
    T-1 SJSYE%MD%.DBF trans to EXCEL
    :param dbf_path: DBF file abpath
    :param excel_path:  DBF file trans to excel  file abpath
    """
    data = DBF(sjsye_dbf_path, encoding='GBK')
    df = pd.DataFrame(iter(data))
    #print('数据表',df)
    df.to_excel(excel_path, index=False)
    #df_filter=df[df['YEZJZH']=='B001650872'] #正式环境
    df_filter=df[df['YEZJZH']=='B0XXXXX872']
    min_backup_data = df_filter['YEZDBF']  # 最低备付数据
    MiniProvisioYEZDBF=min_backup_data.values[0]
    #print('最低备付数据',MiniProvisioYEZDBF)
    return MiniProvisioYEZDBF

def SJSQS0_get_Data(sjsqs0_dbf_path,excel_path):
    """
    SJSQS0.DBF trans to EXCEL
    :param dbf_path: DBF file abpath
    :param excel_path:  DBF file trans to excel  file abpath
    """
    data = DBF(sjsqs0_dbf_path, encoding='GBK')
    df = pd.DataFrame(iter(data))
    #print('数据表',df)
    #df.to_excel(excel_path, index=False)
    #print('表的行列',df.shape)
    if df.shape==(0,0):
        return [0,0]
    else:
        df_filter=df[(df['QSZJZH']=='B001650872')&(df['QSSJLX']=='TZ')&(df['QSYWLB']=='QSHY')] #正式环境
        #df_filter=df[(df['QSZJZH']=='B0XXXXX872')&(df['QSSJLX']=='TZ')&(df['QSYWLB']=='QSHY')]
        if df_filter.shape[0]!=0:
            qingsuan_data = df_filter['QSSFJE']  # 清算收付金QSSFJE

            LiquidateMoneyQSSFJE=qingsuan_data.values[0]
            #print('清算收付金QSSFJE',LiquidateMoneyQSSFJE )
            pici_data = df_filter['QSBYBZ']  # 批次QSBYBZ
            BatchQSBYBZ=pici_data.values[0]
            #print('批次QSBYBZ',BatchQSBYBZ)
            return [LiquidateMoneyQSSFJE,BatchQSBYBZ]
        else:
            return [0,0]





if __name__ == '__main__':
    excel_path=r"C:\Users\26981\Desktop\台账.xlsx"

    test_time=datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') #检测时间
    #print('检测时间：',test_time)
    #深交所提取数据
    taizhang = r"C:\Users\26981\Desktop\深交所现货席位金.xlsx"
    sjsye_dbf_path=r"C:\Users\26981\Desktop\SJSYE0624.DBF"
    sjsqs0_dbf_path=r"C:\Users\26981\Desktop\SJSQS00627.DBF"
    #sjsqs0_dbf_path=r"C:\Users\26981\Desktop\SJSQS00701.DBF"
    MiniProvisioYEZDBF=SJSYE_get_Data(sjsye_dbf_path, excel_path) #最低备付YEZDBF
    LiquidateMoneyQSSFJE,BatchQSBYBZ=SJSQS0_get_Data(sjsqs0_dbf_path, excel_path)
    #print('888888888',LiquidateMoneyQSSFJE,BatchQSBYBZ)

    df = pd.read_excel(taizhang)
    df_shape = df.shape
    #print('台账行列', df_shape)
    row = df_shape[0]
    #print('目前行数：', row)
    loc = 'A' + str(row + 2)
    #print('写入位置：', loc)
    FundAccountYEZJZH= 'B001650872' #资金账号YEZJZH
    WithdrawableMoney='' #可提款金额
    Result=0 #结果
    #Result= WithdrawableMoney+LiquidateMoneyQSSFJE-MiniProvisioYEZDBF#结果
    JudgeResult='' #判断结果
    '''
    if Result<0:
        JudgeResult='risk'
    else:
        JudgeResult='normal'
    '''
    PredictIncome='' #预计入金情况（万）
    ActualIncome='' #实际入金（万）
    ActualWithdraw='' #实际出金（万）
    data=[test_time,FundAccountYEZJZH,WithdrawableMoney,MiniProvisioYEZDBF,LiquidateMoneyQSSFJE,BatchQSBYBZ,Result,JudgeResult,PredictIncome,ActualIncome,ActualWithdraw]
    if JudgeResult=='' and BatchQSBYBZ==3:
        print(True,',',test_time,',',FundAccountYEZJZH,',',WithdrawableMoney,',',MiniProvisioYEZDBF,',',LiquidateMoneyQSSFJE,',',BatchQSBYBZ,',',Result,',',JudgeResult,',',PredictIncome,',',ActualIncome,',','null')
        write_row(taizhang,'sheet1',loc,data)
    else:
        print(False)


