import requests
import xlwt
import re
from bs4 import BeautifulSoup

address = input('请输入网址:')

# 从输入的网址中查找关键字，以便通过api访问
def find_parameter(address_input, parameter):
    try:
        match = re.search(parameter+'=', address_input)
    except:
        print('WRONG！请检查网址是否正确或联系开发人员雾路\n微信号：o8s16se34')
        input('按任意键结束')
        exit()
    for i in range(match.end(), len(address_input)):
        if address_input[i] == '&':
            return address_input[match.end():i]
    return address_input[match.end():]


bid = find_parameter(address, 'bid')
tid = find_parameter(address, 'tid')
url = 'https://chexie.net/api/jiekouapi.php'
data_in = {
    'ask': 'show',
    'ip': '0',
    'token': '0',
    'bid': bid,
    'tid': tid,
}
r = requests.Session().post(url='https://chexie.net/api/jiekouapi.php',data=data_in)

soup = BeautifulSoup(r.text.replace('<text><![CDATA[','<text><![CDATA[<div></div>'), 'lxml')

# 创建Floor类。用于储存楼层信息
class Floor:
    def __init__(self, index, id, content, condition):
        self.floor_index = index  # 楼层索引
        self.id = id  # 会员ID，用于对照ID|姓名
        self.content = content  # 该楼层报名帖内容
        self.condition = condition  # 是否已经删楼

i=0
j=0
floor_list = list()
floor_list.append(Floor(index=i, id='0',content='',condition=False))

# 使用BS4提取楼层信息
for element in soup.recursiveChildGenerator():
    if element.name in ['pid']:
        floor_list[i].content = floor_list[i].content.replace(']]>','')
        floor_list[i].content = floor_list[i].content.split('\n')
        i=i+1
        floor_list.append(Floor(index=i, id='0',content=str(),condition=True))
        floor_list[i].content += ('p='+ str(i+1) + '\n')

    if isinstance(element, str):
        floor_list[i].content += element.strip()  # 添加文本内容并去除首尾空白
    if element.name in ['br', 'div', 'p']:
        floor_list[i].content += '\n'  # 遇到指定标签时添加换行符
    if element.name in ['strike']:
        floor_list[i].condition = False

floor_list[i].content = floor_list[i].content.replace(']]>','')
floor_list[i].content = floor_list[i].content.split('\n')

floor_list = floor_list[2:]

# 用共同元素数量占模式字符串比例来判断是否为所需的类型
def calculate_similarity(pattern, text):
    set_pattern = set(pattern)
    set_text = set(text)
    intersection = len(set_pattern&set_text)
    return intersection/len(set_pattern)

# 通过对比确定报名帖的各个key（无需再手动输入）
def select_key(floor_list, n=int(len(floor_list)/3), accuracy=0.5):
    text_before_first_colon = dict()
    for floor in floor_list[1:n]:
        for element in floor.content:
            try:
                first_colon = re.search(pattern=":|：", string=element)
                temp = element[:first_colon.end()]
                if temp in text_before_first_colon.keys():
                    text_before_first_colon[temp] = text_before_first_colon[temp] + 1
                else:
                    text_before_first_colon[temp] = 1
            except:
                continue
    print(text_before_first_colon)
    global key
    key = list()
    for k in text_before_first_colon.keys():
        if text_before_first_colon[k] > accuracy*n:
            key.append(k)
    return key


key = select_key(floor_list)
print(key)

# 创建输出
workbook = xlwt.Workbook()
workbook.data_only = True
namelist = workbook.add_sheet('Sheet')

namelist.write(0, 0, '序号')
for i, elem in enumerate(key):
    namelist.write(0, i+1, elem)
namelist.write(0, len(key)+1, '信息缺失情况')
namelist.write(0, len(key)+2, '会员更改信息和删楼情况')

# 输出成excel表格
for floor in floor_list:
    namelist.write(floor.floor_index, 0, floor.floor_index)
    print(floor.content)
    if floor.condition == False:
        namelist.write(floor.floor_index, len(key)+2, '该会员可能更改过信息或已经已经删楼！！！')
        print(floor.floor_index,'该会员可能更改过信息或已经已经删楼！！！')
    for i, k in enumerate(key):
        judge = 0
        for elem in floor.content:
            #print(k, elem)
            #print(floor.floor_index)
            if calculate_similarity(pattern=k, text=elem) > 0.5:
                namelist.write(floor.floor_index, i+1, elem[len(k):])
                floor.content.remove(elem)
                judge = 1
                break
        if judge == 0:
            try:
                namelist.write(floor.floor_index, len(key) + 1, '信息缺失')
                print(str(floor.floor_index), '信息缺失')
            except:
                pass

# 保存
for i in range(255):
    try:
        workbook.save('namelist'+ str(i) +'.xls')
        print('已保存为'+'namelist'+ str(i) +'.xls')
        break
    except:
        i = i + 1

input('输入任意键退出。感谢你使用由<雾路|曹三省>制作的拉练报名信息统计工具。'
      '\n感谢参与测试人员<余割|陈少春><彤彤|侯依彤>\n如有疑问请联系微信c602\n微信号：o8s16se34')
