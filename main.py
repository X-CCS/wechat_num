# pip3 install xlrd xlwt xlutils
# pip install -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com xlrd xlwt xlutils
import json
import codecs
# # 导入 xlwt 库
import xlwt
 
json_filename = '/Users/ccs/Desktop/myRepo/wechat_num/json_data/亿数通三周年分享VIP11群.json' #这是json文件存放的位置
excel_filename = './excel_data/wechat_w.xls'   #这是保存的excel文件的位置
# 创建 xls 文件对象
wb = xlwt.Workbook()

# 新增两个表单页
sh1 = wb.add_sheet('加人微信号')
# sh2 = wb.add_sheet('汇总')


# 然后按照位置来添加数据,第一个参数是行，第二个参数是列
# 写入第一个sheet
# 表头
sh1.write(0, 0, '微信名')
sh1.write(0, 1, '微信号')

# excel_file=open(excel_filename,'w')
with open(json_filename) as f:
     pop_data = json.load(codecs.open(json_filename, 'r+', 'utf-8-sig'))
    #  pop_data = json.load(f)
     print(pop_data)
     for i,pop_dict in enumerate(pop_data):
        #  i.sort()
        num = i+1
        # print(num)
        # print(i.sort())
        nick_name = pop_dict['nick_name']
        # print("nick_name:",nick_name)
        wxid = pop_dict['wxid']
        # print("wxid:",wxid)
        # 写入内容
        sh1.write(num, 0, nick_name) # 微信名
        sh1.write(num, 1,wxid) # 微信号码

# 最后保存文件即可
wb.save(excel_filename)
