import xlrd
import xlwt
import wx
import random
import time
import _thread

def open_file(ecents):
    global file_path, run_list, all_list, list_title
    list_title = []
    run_list = []
    all_list = []
    wildcard = "Excel 工作簿 (*.xlsx)|*.xlsx|""Excel 97-2003 工作簿 (*.xls)|*.xls|" "All files (*.*)|*.*"
    fileDialog = wx.FileDialog(panel, message ="请选择数据文件", wildcard = wildcard, style = wx.FD_OPEN)
    dialogResult = fileDialog.ShowModal()
    if dialogResult ==  wx.ID_OK:
        file_path = fileDialog.GetPath()
        first_empty_text.SetLabel("打开的文件路径是："+file_path)
        fileDialog.Destroy()
    work_book = xlrd.open_workbook(file_path)
    sheet1 = work_book.sheet_by_index(0)
    number_of_rows = sheet1.nrows
    number_of_columns = sheet1.ncols
    list_title = sheet1.row_values(0)
    for row in range(1,number_of_rows):
        all_list.append(sheet1.row_values(row))
        if type(sheet1.row(row)[0].value) != float:
            run_list.append(str(sheet1.row(row)[0].value))
        else:
            run_list.append(str(int(sheet1.row(row)[0].value)))

def parameter_configuration(events):
    global run_number
    dlg = wx.TextEntryDialog(None, "要抽取的名额是：","参数配置")  
    if dlg.ShowModal() == wx.ID_OK:      
        run_number = int(dlg.GetValue())

def start_run(events):
    global going,current_number,run_number
    if current_number != run_number:
        if going == True:
            pass
        else:
            going = True
            _thread.start_new_thread(random_list,())
    else:
        process_text.SetLabel("抽取完成！")

def save_file(events):
    global final_list,all_list,result_list,list_title
    result_list.append(list_title)
    result_book = xlwt.Workbook(encoding = 'ascii')
    result_sheet = result_book.add_sheet('抽取结果')
    for r in range(len(all_list)):
        for m in range(len(final_list)):
            if type(all_list[r][0]) != float:
                s = str(all_list[r][0])
                if final_list[m] == s:
                    result_list.append(all_list[r])
                else:
                    continue
            else:
                s = str(int(all_list[r][0]))
                if final_list[m] == s:
                    result_list.append(all_list[r])
                else:
                    continue
    for r in range(len(result_list)):
        for i in range (len(result_list[r])):
            result_sheet.write(r,i,result_list[r][i])
    localtime = time.strftime("%Y%m%d%H%M%S", time.localtime())
    dlg = wx.TextEntryDialog(None, "要保存到文件名为：","结果保存")  
    if dlg.ShowModal() == wx.ID_OK:      
        save_file_name = str(dlg.GetValue())+".xls"
    result_book.save("【"+str(localtime)+"】"+save_file_name)
    
def stop_run(events):
    global going
    going = False

def random_list():
    global current_choice,current_number,run_list,going,final_list
    while True:
        current_choice = str(random.choice(run_list))
        process_text.SetLabel(current_choice)
        time.sleep(0.01)
        if going == False:
            current_number = current_number + 1
            run_list.remove(current_choice)
            final_list.append(current_choice)
            second_empty_text.SetLabel("恭喜  "+current_choice+"  ！！！")
            break

def README_Text(events):
    README_Frame = wx.Frame(None,title = "使用文档",pos = (800,400),size = (500,400))
    README_panel = wx.Panel(README_Frame)

    README_TITLE_Text = wx.StaticText(README_panel, id=wx.ID_ANY, label= "使用文档", style = wx.ALIGN_LEFT|wx.ST_NO_AUTORESIZE)
    README_TITLE_Text.SetFont(wx.Font(pointSize = 20, family = wx.MODERN, style = wx.NORMAL, weight = wx.BOLD, faceName = "Microsoft YaHei"))
    README_TITLE_Text.SetForegroundColour('black')
    ALL_Text_label = ("1、本程序只适用于Excel类型文件导入,如需其他格式文件请自行修改源码。\n2、程序默认设置Excel文件中第一行为标题行，设置第一列为显示列\n3、保存文件使用了xlwt模块，受模块限制，只能保存为xls格式，请不要选择过多\n的数据，否则可能导致写入出错\n4、使用顺序为\n①导入文件\n②配置参数\n③点击“开始抽取”屏幕机即会开始滚动\n④点击“停止抽取”滚动即停止，结果会自动储存\n⑤当抽取次数达到上限是会弹出提醒“抽取完成！\n5、“保存文件”选项会将结果导出为xls文件保存在软件同目录下")
    README_ALL_Text = wx.StaticText(README_panel, id=wx.ID_ANY, label= ALL_Text_label, style = wx.ALIGN_LEFT|wx.ST_ELLIPSIZE_END|wx.ST_NO_AUTORESIZE)
    README_ALL_Text.SetFont(wx.Font(pointSize = 10, family = wx.MODERN, style = wx.NORMAL, weight = wx.BOLD, faceName = "Microsoft YaHei"))
    README_ALL_Text.SetForegroundColour('black')

    README_F = wx.BoxSizer(wx.VERTICAL)
    README_F.Add(README_TITLE_Text, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)
    README_F.Add(README_ALL_Text, proportion = 4, flag = wx.EXPAND|wx.ALL, border = 3)
    README_panel.SetSizer(README_F)
    README_Frame.Show()





#相关变量定义
run_number = 0
current_number = 0
current_choice = ""
going = False
i = 0
file_path = ""
list_title = []
run_list = []
final_list = []
all_list = []
result_list = []

#GUI绘制
app = wx.App()
frame = wx.Frame(None,title = "考研自习室抽取",pos = (1000,200),size = (500,400))
 
panel = wx.Panel(frame)

open_button = wx.Button(panel,label = "打开文件")
open_button.Bind(wx.EVT_BUTTON,open_file)

save_button = wx.Button(panel,label = "保存文件")
save_button.Bind(wx.EVT_BUTTON,save_file)

parameter_configuration_button = wx.Button(panel,label = "参数配置")
parameter_configuration_button.Bind(wx.EVT_BUTTON,parameter_configuration)

README_Text_button = wx.Button(panel,label = "使用说明")
README_Text_button.Bind(wx.EVT_BUTTON,README_Text)

start_run_button = wx.Button(panel,label = "开始抽取")
start_run_button.Bind(wx.EVT_BUTTON,start_run)

stop_run_button = wx.Button(panel, label = "停止抽取")
stop_run_button.Bind(wx.EVT_BUTTON,stop_run)


process_text= wx.StaticText(panel, id=wx.ID_ANY, label= "颜亦童", style = wx.ALIGN_CENTRE_HORIZONTAL|wx.ST_ELLIPSIZE_END|wx.ST_NO_AUTORESIZE)
process_text.SetFont(wx.Font(pointSize = 40, family = wx.MODERN, style = wx.NORMAL, weight = wx.BOLD, faceName = "Microsoft YaHei"))
process_text.SetForegroundColour('purple')

first_empty_text= wx.StaticText(panel, id=wx.ID_ANY, label="",style = wx.ALIGN_CENTRE_HORIZONTAL|wx.ST_ELLIPSIZE_END|wx.ST_NO_AUTORESIZE)
second_empty_text= wx.StaticText(panel, id=wx.ID_ANY, label="本程序只适用于Excel类型文件导入\n使用有疑问请点击“使用说明”查看使用文档",style = wx.ALIGN_CENTRE_HORIZONTAL|wx.ST_ELLIPSIZE_END|wx.ST_NO_AUTORESIZE)
second_empty_text.SetFont(wx.Font(pointSize = 10, family = wx.MODERN, style = wx.NORMAL, weight = wx.BOLD, faceName = "Microsoft YaHei"))
second_empty_text.SetForegroundColour('red')
third_empty_text= wx.StaticText(panel, id=wx.ID_ANY, label="",style = wx.ALIGN_CENTRE_HORIZONTAL|wx.ST_ELLIPSIZE_END|wx.ST_NO_AUTORESIZE)
fourth_empty_text= wx.StaticText(panel, id=wx.ID_ANY, label="",style = wx.ALIGN_CENTRE_HORIZONTAL|wx.ST_ELLIPSIZE_END|wx.ST_NO_AUTORESIZE)

menu_box = wx.BoxSizer()
menu_box.Add(open_button, proportion = 1, flag = wx.EXPAND|wx.ALL,border = 3)
menu_box.Add(save_button, proportion = 1, flag = wx.EXPAND|wx.ALL,border = 3)
menu_box.Add(parameter_configuration_button, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)
menu_box.Add(README_Text_button, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)

run_box = wx.BoxSizer()
run_box.Add(start_run_button, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 30)
run_box.Add(stop_run_button, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 30)

v_box = wx.BoxSizer(wx.VERTICAL)
v_box.Add(menu_box, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)
v_box.Add(first_empty_text, proportion = 2, flag = wx.EXPAND|wx.ALL, border = 3)
v_box.Add(process_text, proportion = 4, flag = wx.EXPAND|wx.ALL, border = 3)
v_box.Add(second_empty_text, proportion = 2, flag = wx.EXPAND|wx.ALL, border = 3)
v_box.Add(run_box, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 3)

panel.SetSizer(v_box)

frame.Show()
app.MainLoop()
