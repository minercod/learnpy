import calendar
import xlwt
import xlrd
import random

#人员类
class Person(object):
    def __init__(self,name,totalcount,nooncount,evecount,weekdaycount):
        self.name         = name
        self.totalcount   = totalcount
        self.nooncount    = nooncount
        self.evecount     = evecount
        self.weekday      = weekday


#判断值班当天前后星期，用于判断中班、夜班
def judgeweek(Person_weekday,weekday):
    week_condition = {'一': ('日','一','二'),
                      '二': ('一','二','三'),
                      '三': ('二','三','四'),
                      '四': ('三','四','五'),
                      '五': ('四','五','六'),
                      '六': ('五','六','日'),
                      '日': ('六','日','一')}
    for key,value in week_condition:
        if Person_weekday == key:
            condition = value
    if weekday in condition:
        return False
    else:
        return True

#判断值班当天星期，用于判断早班
def judgetoday(Person_weekday,weekday):
        if Person_weekday == weekday:
            return False
        else:
            return True


#判断闰年
def initial_months(year):
    months = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if calendar.isleap(year):
        months[1] = 29
    else:
        months[1] = 28
    return months

#根据年份月份判断当月天数
def initial_days(year,month):
    months = initial_months(year)
    days   = months[month-1]
    return days

#返回日期的星期
def weekday(year,month,day):
    weekdaynum = calendar.weekday(year,month,day)
    weekdays   = ['一','二','三','四','五','六','日']
    weekday    = weekdays[weekdaynum]
    return weekday

#获取年份
def get_year():
    while True:
        year = int(input('请输入年份（注意：请输入数字，2000～2100）：'))
        if year > 1999 and year < 2100:
            break
    return year

#获取月份
def get_month():
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
    for i in range(1,sheet.nrows):
        row_value = sheet.row_values(i)
        if row_value[0] != '':
            x = Person(row_value[0],int(row_value[1]),int(row_value[2]),int(row_value[3]),row_value[4])
            profile.append(x)
            del x
    return profile

#初始化数据容器
def initial_list(year,month,days):
    dutylist = []
    daylist   = ['0','0','0','0','0']
    for i in range(days):
        daylist[0] = i+1
        dutylist.append(daylist.copy())
    for i in range(days):
        dutylist[i][1] = weekday(year,month,i+1)
    return dutylist

#开始排班
def chooseduty(profile,dutylist):
    #夜班
    for j in range(len(dutylist)):
        x = random.randint(0,len(profile))
        dutylist[j][4]=profile[x].name









#初始化输出表格格式
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

def write_to_excel(year,month,dutylist):
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
    for i in range(len(dutylist)):
        new_sheet.write(i+3,0,str(dutylist[i][0]),set_style('Times New Roman',300))
        new_sheet.write(i+3,1,dutylist[i][1],set_style('宋体',300))
        new_sheet.write(i+3,2,dutylist[i][2],set_style('宋体',300))
        new_sheet.write(i+3,3,dutylist[i][3],set_style('宋体',300))
        new_sheet.write(i+3,4,dutylist[i][4],set_style('宋体',300))

    #设置列宽
    for l in range(2):
        new_sheet.col(l).width = 256*8
    for i in range(2,5):
        new_sheet.col(i).width = 256*25

    #设置行高
    new_sheet.row(0).height_mismatch = True
    new_sheet.row(0).height          = 800

    new_workbook.save(str(year)+'.'+str(month)+'dutylist.xls')


def main():
    year      = get_year()
    month     = get_month()
    days      = initial_days(year,month)
    dutylist = initial_list(year,month,days)
    write_to_excel(year,month,dutylist)
    print("%s年%s月有%s天"%(year,month,days))



if __name__ == '__main__':
    #main()
    print(get_profile())








