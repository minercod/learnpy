import calendar
import xlwt,xlrd
import random

#欢迎语
def intro():
    print('------------------------------------------自动化带班表生成程序-----------------------------------------')
    print('   注意: 程序运行需先自行设定main()函数中的corp_name;可利用except_check()函数来单独排除某人不带班的日期   ')
    print('------------------------------------------------------------------------------------------------------')

#人员类
class Person(object):
    def __init__(self, name, morningcount, nooncount, evecount, weekday):
        self.name = name
        self.morningcount = morningcount
        self.nooncount = nooncount
        self.evecount = evecount
        self.weekday = weekday

#判断值班当天前后星期，用于判断中班、夜班
def judgeweek(person_weekday, weekday):
    week_condition = {'一': ('日', '二'),
                      '二': ('一', '三'),
                      '三': ('二', '四'),
                      '四': ('三', '五'),
                      '五': ('四', '六'),
                      '六': ('五', '日'),
                      '日': ('六', '一')}
    if person_weekday in week_condition:
        conditions = week_condition[person_weekday]
        if weekday in conditions:
            return False
        else:
            return True
    else:
        return True

#判断值班当天星期，用于判断早班
def judgetoday(person_weekday, weekday):
        if person_weekday == weekday:
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
def initial_days(year, month):
    months = initial_months(year)
    days = months[month-1]
    return days

#返回日期的星期
def weekday(year, month, day):
    weekdaynum = calendar.weekday(year, month, day)
    weekdays = ['一', '二', '三', '四', '五', '六', '日']
    weekday = weekdays[weekdaynum]
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
    profile = []
    xls = xlrd.open_workbook('profile.xls')
    sheet = xls.sheets()[0]
    for i in range(1, sheet.nrows):
        row_value = sheet.row_values(i)
        if row_value[0] != '':
            x = Person(row_value[0], int(row_value[1]), int(row_value[2]), int(row_value[3]), row_value[4])
            profile.append(x)
            del x
    return profile

#初始化数据容器
def initial_list(year, month, days):
    dutylist = []
    daylist = ['0', '0', '0', '0', '0']
    for i in range(days):
        daylist[0] = i+1
        dutylist.append(daylist.copy())
    for i in range(days):
        dutylist[i][1] = weekday(year, month, i+1)
    return dutylist

#判断重复
def judgerepeat(name, dutylist, j):
    if j == 0:
        if name in dutylist[0]:
            return False
        else:
            return True
    elif j == 1:
        if name in dutylist[1] or name in dutylist[0]:
            return False
        else:
            return True
    else:
        if name in dutylist[j] or name in dutylist[j-1] or name in dutylist[j-2]:
            return False
        else:
            return True

#排除检查
def except_check(name, j):
    if name == '   ':
        if j in (8, 9, 18, 19, 28, 29, 30):
            return False
        else:
            return True
    elif name == '    ':
        if j in (28, 29, 30):
            return False
        else:
            return True
    elif name == ' ':
        if j in (9, 19, 29, 30):
            return False
        else:
            return True
    else:
        return True

#开始排班
def chooseduty(profile, dutylist):
    len_pro = len(profile)

    for j in range(len(dutylist)):

        #夜班
        while True:
            x = random.randint(0, len_pro-1)
            if profile[x].evecount != 0 and judgeweek(profile[x].weekday, dutylist[j][1]) and judgerepeat(profile[x].name, dutylist, j) and except_check(profile[x].name, j):
                break
        dutylist[j][4] = profile[x].name
        profile[x].evecount -= 1

        #中班
        while True:
            y = random.randint(0, len_pro-1)
            if profile[y].nooncount != 0 and judgeweek(profile[y].weekday, dutylist[j][1]) and judgerepeat(profile[y].name, dutylist, j) and except_check(profile[x].name, j):
                break
        dutylist[j][3] = profile[y].name
        profile[y].nooncount -= 1

        #早班
        while True:
            z = random.randint(0, len_pro-1)
            if profile[z].morningcount != 0 and judgetoday(profile[z].weekday, dutylist[j][1]) and judgerepeat(profile[z].name, dutylist, j) and except_check(profile[x].name, j):
                break
        dutylist[j][2] = profile[z].name
        profile[z].morningcount -= 1
    print('……排班表已生成，正在写入xls文件……')
    return dutylist

#初始化输出表格格式
def set_style(font_name, font_height, bold = False, bordersset = False):
    style = xlwt.XFStyle()

    #字体设置
    font = xlwt.Font()
    font.name = font_name
    font.height = font_height  #20=1pt
    font.bold = bold

    #单元格设置
    borders = xlwt.Borders()
    if bordersset:
        borders.left = 0
        borders.right = 0
        borders.top = 0
        borders.bottom = 0
    else:
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN

    #居中设置
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER

    style.font = font
    style.borders = borders
    style.alignment = alignment

    return style

def write_to_excel(year, month, dutylist, corp_name):
    new_workbook = xlwt.Workbook()
    new_sheet    = new_workbook.add_sheet('带班')

    #初始化表头
    new_sheet.write_merge(0, 0, 0, 4, corp_name+'领导带班表['+str(year)+'年'+str(month)+'月]', set_style('宋体', 400, True, True))
    new_sheet.write_merge(1, 2, 0, 0, '日期', set_style('宋体', 300))
    new_sheet.write_merge(1, 2, 1, 1, '星期', set_style('宋体', 300))
    new_sheet.write_merge(1, 1, 2, 4, '带班人员', set_style('宋体', 300))
    new_sheet.write(2, 2, '早班', set_style('宋体', 300))
    new_sheet.write(2, 3, '中班', set_style('宋体', 300))
    new_sheet.write(2, 4, '夜班', set_style('宋体', 300))

    #充填数据
    for i in range(len(dutylist)):
        new_sheet.write(i+3, 0, str(dutylist[i][0]), set_style('Times New Roman', 300))
        new_sheet.write(i+3, 1, dutylist[i][1], set_style('宋体', 300))
        new_sheet.write(i+3, 2, dutylist[i][2], set_style('宋体', 300))
        new_sheet.write(i+3, 3, dutylist[i][3], set_style('宋体', 300))
        new_sheet.write(i+3, 4, dutylist[i][4], set_style('宋体', 300))

    #设置列宽
    for l in range(2):
        new_sheet.col(l).width = 256*8
    for i in range(2, 5):
        new_sheet.col(i).width = 256*25

    #设置行高
    new_sheet.row(0).height_mismatch = True
    new_sheet.row(0).height = 800

    new_workbook.save(str(year)+'.'+str(month)+'dutylist.xls')

def main():
    corp_name = '    '
    intro()
    year = get_year()
    month = get_month()
    days = initial_days(year, month)
    print("%s年%s月有%s天" % (year, month, days))
    dutylistex = initial_list(year, month, days)
    profile = get_profile()
    dutylist = chooseduty(profile, dutylistex)
    write_to_excel(year, month, dutylist, corp_name)
    print('……文件已写入完毕……')

if __name__ == '__main__':
    main()