import pandas as pd
import xlrd as xl

def read(c, ws):
    for (int i = 0;i<14;i++):


def main():

    fm=pd.read_excel(r"F:\aSSQ\ssq.xlsm", sheet_name="data")
    wk=xl.open_workbook(r"F:\aSSQ\ssq.xlsm")
    ws=wk.sheet_by_name("data")
    #用该方法读取表格和表单里的单元格的数据
    rows = ws.nrows()
    cols = ws.ncols()
    read(cols, ws)

if __name__ == '__main__':
    main()