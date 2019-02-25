import xlrd
import xlsxwriter
import sys
import os

# 获取当前目录下的考勤表excel文件
def getExcelFile():
    files = os.listdir(sys.path[0])
    for file in files:
        str1 = os.path.splitext(file)[0]
        str2 = os.path.splitext(file)[1]
        if str2 == '.xlsx' and "全体成员考勤表" in str1:
            return file

# 获取excel信息
def getExcelInfoList():
    x1 = xlrd.open_workbook(file_complete_path)
    sheet = x1.sheet_by_name("月度考勤表")
    # 获取名字列表
    value_name_list = []
    name_list = sheet.col_values(1)
    for name in name_list:
        if name != "" and name != "姓名":
            # print(name)
            value_name_list.append(name)

    # 下班时间列 的列表
    row_list = []
    first_row = 4
    print("人数" + str(len(value_name_list)))
    for value in value_name_list:
        off_time_list = []
        for cell in sheet.row(first_row):
            if ":" in cell.value:
                off_time_list.append(cell.value)
        row_list.append(off_time_list)
        first_row = first_row + 4

    return value_name_list, row_list

# 写数据入excel
def writeExcel(name_list, row_list):
    x2 = xlsxwriter.Workbook("月度加班时间统计.xlsx")
    sheet_1 = x2.add_worksheet("月度加班时间总计")
    # sheet_1.set_column(0, None, x2.add_format({'bold': True}))
    # sheet_1.set_column(1, None, None)
    sheet_1.write_string("A1", "名字")
    sheet_1.write_string("B1", "加班时间总长")

    sheet_1.write_column("A2", name_list)
    over_time_list = []
    for row_data in row_list:
        total = 0
        print(name_list[row_list.index(row_data)])
        for value_time in row_data:
            print(value_time)
            value1 = 0
            value2 = 0
            value_time_list = str.split(value_time, ":")
            hours = int(value_time_list[0])
            minutes = int(value_time_list[1])
            if hours > 20:
                value1 = hours - 20
                if 20 <= minutes < 50:
                    value2 = 0.5
                elif minutes >= 50:
                    value2 = 1
            total = total + value1 + value2
            print("时间：" + str(value1 + value2))
        print("总时间：" + str(total))
        over_time_list.append(total)

    sheet_1.write_column("B2", over_time_list)
    x2.close()
    pass

# 主程序入口
def main():
    name_list, row_list = getExcelInfoList()
    writeExcel(name_list, row_list)

if __name__ == "__main__":
    # excel
    excel_file = getExcelFile()
    # 根目录
    root_path = sys.path[0]
    # 文件完整路径
    file_complete_path = os.path.join(root_path, excel_file)
    main()