import xlrd
import xlwt
import pysnooper

from excel_reorder import re_order
@pysnooper.snoop()
def excel_read(fil_name):
    data_rd = xlrd.open_workbook(fil_name)
    table_rd = data_rd.sheet_by_index(0)
    nrows = table_rd.nrows
    ncols = table_rd.ncols


    data_write = xlwt.Workbook()
    table_wt = data_write.add_sheet('first',cell_overwrite_ok = True)

    wr_rows = 0
    wl_rows = 0

    change_label = True
    for i_col in range(ncols):
        for i_row in range(nrows):
            if i_col < 3:
                continue
            testdata = table_rd.col(i_col)[i_row].value

            if testdata == "数量合计":
                break

            if testdata is '':
                continue

            if not isinstance(testdata, str):
                label = table_rd.col(1)[i_row].value
                units = table_rd.col(2)[i_row].value
                table_wt.write(wr_rows, 1, label)
                table_wt.write(wr_rows, 2, units)
                table_wt.write(wr_rows, 3, testdata)
            else:
                if wr_rows is not 0:
                    wr_rows += 1
                table_wt.write(wr_rows, 0, testdata)
                table_wt.write(wr_rows, 1, '品名')
                table_wt.write(wr_rows, 2, '单位')
                table_wt.write(wr_rows, 3, '数量')

            wr_rows += 1

    name = "output.xls"
    data_write.save(name)

    name = re_order(name)
    return name

if __name__ == '__main__':
    n = excel_read('干货.xls')
    print(n)
