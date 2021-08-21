from PyQt5.QtWidgets import QMessageBox
import xlwings as xw
from id_validator import validator

"""
    适用于33类人的excel验证程序
    作者:@8bit

"""


def IsChinese(string):
    if string:
        for chart in string:
            if not '\u4e00' <= chart <= '\u9fa5':
                return False
        return True

def IsNumber(number):
     number = str(number)
     local = number.find(".")
     if local != -1:
         number = number[0:local]
     if len(number) == 11:
         return True
     return False


def openexcel(url):
    global wb,sht #打开工作簿
    wb = xw.Book(url) #读取数据
    sht = wb.sheets["sheet1"] #获取总行数
    global count_row
    count_row = sht.range("a1").expand("table").rows.count
    location = "A4:K{}".format(count_row)
    range_date = sht.range(location).value
    return range_date


def mid(str,startpos,length=0): #length为可选参数，若不写则默认值为0
    if length == 0:
       length = len(str) - startpos + 1 #判断length是否为默认值0，若是，则计算实际长度
    endpos = startpos + length - 1
    mid = str[startpos - 1:endpos]
    return mid


def InputErrorRange(row,error):
    error_name = ""
    error = "000" + mid(bin(error),3)
    for i in range(0,4):
        if error[-(i + 1)] == "1":
            error_name = error_name + error_string[i]
    if error_name != "":
        sht.range("L{}".format(row)).value = error_name
    else:
        sht.range("L{}".format(row)).value = "合格"



def verify(list):
    errorlog = 0
    #if not IsChinese(list[0]):
    #    errorlog = errorlog + 1
    if not validator.is_valid(str(list[1])):
        errorlog = errorlog + 2
    if not IsNumber(list[2]):
        errorlog = errorlog + 4
    if list[6] not in people_type:
        errorlog = errorlog + 8
    return errorlog


def main(user_input):
    #user_input = input("请输入要执行的excel位置:")
    #打开并读取数据
    range_date = openexcel(user_input)
    #L列先写入正确信息
    sht.range("L4:L{}".format(count_row)).value = "已读取"
    #验证
    global error_string,people_type
    error_string = ["姓名错误;","身份证错误;","联系方式错误;","人员分类不在列表中;"]
    people_type = [ '密切接触者、次密切接触者',
                    '次次密切接触者',
                    '境外入冀返冀人员',
                    '高、中风险地区来冀返冀人员', 
                    '高、中风险地区所在县（市、区）和出现阳性感染者所在县（市、区）返回人员',
                    '隔离场所管理和服务人员',
                    '进口冷链食品监管和从业人员',
                    '口岸检疫和边防检查人员',
                    '口岸直接接触进口货物从业人员',
                    '国际交通运输工具从业人员',
                    '船舶引航员等登临外籍船舶作业人员',
                    '监狱人员',
                    '监所工作人员及新收被监管人员（公安、司法）',
                    '戒毒所人员（公安、司法）',
                    '各级医疗卫生机构工作人员',
                    '发热门诊患者',
                    '新住院患者及陪护人员',
                    '各类交通运输服务保障人员',
                    '社会福利养老机构工作人员',
                    '市场监管系统一线工作人员',
                    '旅游景区及配套服务设施工作人员',
                    '宾馆饭店工作人员',
                    '影剧院及娱乐场所工作人员',
                    '商场超市工作人员',
                    '建筑工地施工和管理人员',
                    '寄宿制中小学校人员',
                    '大中专院校人员',
                    '各级党政群机关人员'
                  ]
    for first_date in range(0,count_row - 3):
        error_number = verify(range_date[first_date])
        InputErrorRange(first_date + 4,error_number)

if __name__ == "__main__":
    main()
