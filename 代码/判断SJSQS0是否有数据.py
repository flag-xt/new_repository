from dbfread import DBF
import sys
import pandas as pd



def DBF_to_Excel(dbf_path):
    """
    DBF trans to EXCEL
    :param dbf_path: DBF file abpath
    :param excel_path:  DBF file trans to excel  file abpath
    """
    data = DBF(dbf_path, encoding='GBK')
    df = pd.DataFrame(iter(data))

    return df.shape==(0,0)



if __name__ == '__main__':
    #dbf_path = sys.argv[1]
    #dbf_path = r"C:\Users\26981\Desktop\货银\数据\SJSQS00701.DBF"
    dbf_path=r"C:\Users\26981\Desktop\SJSQS00627.DBF"


    df=DBF_to_Excel(dbf_path)
    print(df)
