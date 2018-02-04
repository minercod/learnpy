import calendar
import xlwt
import xlrd

#判断闰年
def initialmonths(year):
    months = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if calendar.isleap(year):
        months[1] = 29
    else:
        months[1] = 28
    return months

#根据年份月份判断当月天数
def initialdays(year,month):
    months = initialmonths(year)
    days   = months[month-1]
    return days

#返回日期的星期
def weekday(year,month,day):
    weekdaynum = calendar.weekday(year,month,day)
    weekdays   = ['一','二','三','四','五','六','日']
    weekday    = weekdays[weekdaynum]
    return weekday

#获取年份
def getyear():
    while True:
        year = int(input('请输入年份（注意：请输入数字，2000～2100）：'))
        if year > 1999 and year < 2100:
            break
    return year

#获取月份
def getmonth():
    while True:
        month = int(input('请输入月份（注意：请输入数字，1～12）：'))
        if month > 0 and month < 13:
            break
    return month

#提取profile
def get_profile():
    profile  = []
    xls      = xlrd.open_workbook('profile.xls')
    sheet    = xls.sheets()[0]
    for i in range(1,15):
        profile.append(sheet.row_values(i))
    for j in range(len(profile)):
        profile[j][1] = int(profile[j][1])
        profile[j][2] = int(profile[j][2])
        profile[j][3] = int(profile[j][3])
    return profile

#初始化数据容器
def initlist(year,month,days):
    monthlist = []
    daylist   = ['0','0','0','0','0']
    for i in range(days):
        daylist[0] = i+1
        monthlist.append(daylist.copy())
    for i in range(days):
        monthlist[i][1] = weekday(year,month,i+1)
    return monthlist

#初始化输出表格
def set_style(font_name,font_height,bold=False,bordersset=False):
    style = xlwt.XFStyle()

    #字体设置
    font = xlwt.Font()
    font.name   = font_name
    font.height = font_height  #20=1pt
    font.bold   = bold

    #单元格设置
    borders = xlwt.Borders()
    if bordersset:
        borders.left   = 0
        borders.right  = 0
        borders.top    = 0
        borders.bottom = 0
    else:
        borders.left   = xlwt.Borders.THIN
        borders.right  = xlwt.Borders.THIN
        borders.top    = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN

    #居中设置
    alignment      = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER

    style.font      = font
    style.borders   = borders
    style.alignment = alignment

    return style

def write_to_excel_xlwt(year,month,monthlist):
    new_workbook = xlwt.Workbook()
    new_sheet    = new_workbook.add_sheet('带班')

    #初始化表头
    new_sheet.write_merge(0,0,0,4,'何家堡煤业领导带班表['+str(year)+'年'+str(month)+'月]',set_style('宋体',400,True,True))
    new_sheet.write_merge(1,2,0,0,'日期',set_style('宋体',300))
    new_sheet.write_merge(1,2,1,1,'星期',set_style('宋体',300))
    new_sheet.write_merge(1,1,2,4,'带班人员',set_style('宋体',300))
    new_sheet.write(2,2,'早班',set_style('宋体',300))
    new_sheet.write(2,3,'中班',set_style('宋体',300))
    new_sheet.write(2,4,'夜班',set_style('宋体',300))

    #充填数据
    for i in range(len(monthlist)):
        new_sheet.write(i+3,0,str(monthlist[i][0]),set_style('Times New Roman',300))
        new_sheet.write(i+3,1,monthlist[i][1],set_style('宋体',300))
        new_sheet.write(i+3,2,monthlist[i][2],set_style('宋体',300))
        new_sheet.write(i+3,3,monthlist[i][3],set_style('宋体',300))
        new_sheet.write(i+3,4,monthlist[i][4],set_style('宋体',300))

    #设置列宽
    for l in range(2):
        new_sheet.col(l).width = 256*8
    for i in range(2,5):
        new_sheet.col(i).width = 256*25

    #设置行高
    new_sheet.row(0).height_mismatch = True
    new_sheet.row(0).height          = 800
    #for j in range(1,len(monthlist)+3):
        #new_sheet.row(i).height_mismatch = True
        #new_sheet.row(i).height = 1000

    new_workbook.save(str(year)+'.'+str(month)+'dutylist.xls')


def main():
    year      = getyear()
    month     = getmonth()
    days      = initialdays(year,month)
    monthlist = initlist(year,month,days)
    write_to_excel_xlwt(year,month,monthlist)
    print("%s年%s月有%s天"%(year,month,days))



if __name__ == '__main__':
    main()
    #x=get_profile()
    #print(x)







