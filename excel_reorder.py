import xlrd
import xlwt
import pysnooper
import datetime
import numpy

@pysnooper.snoop()
def re_order(fil_name):
    """
    Get the input excel file information
    """
    excel_rd = xlrd.open_workbook(fil_name)
    excel_tb = excel_rd.sheet_by_index(0)
    nrows = excel_tb.nrows  # Get the number of origin rows
    ncols = excel_tb.ncols  # Get the number of origin cols

    """
    Create the new excel for output
    """
    excel_wt = xlwt.Workbook()
    table_wt = excel_wt.add_sheet('first', cell_overwrite_ok=True)

    num = int(nrows / 2) # Get the middle rows
    while(1):
        test_data = excel_tb.col(3)[num].value

        """
        Checking the middle 
        Ensure it's empty
        """
        if test_data is not '':
            num -= 1
        else:
            mark_num = num + 1
            break

    """
    Traversing the origin excel file util the middle row
    """
    for i_row in range(nrows):
        for i_col in range(ncols):
            if i_row < mark_num:
                table_wt.write(i_row, i_col, excel_tb.row(i_row)[i_col].value)
            else:   # When Traver to the middle row change to the second
                table_wt.write(i_row - mark_num, i_col+5, excel_tb.row(i_row)[i_col].value)


    today = datetime.date.today()
    now = today.strftime('%m%d')
    name = "{0}.xls".format(now)
    excel_wt.save(name)

    return name



if __name__ == '__main__':
    re_order('0517.xls')