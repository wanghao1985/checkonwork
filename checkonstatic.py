import xlrd

from datetime import date,datetime

def read_excel():
    workbook = xlrd.open_workbook(r'C:\Users\Wanghao\Desktop\kq7.xlsx')   #修改路径和文件名，最好用英文名，中文名未测试
    #print(workbook.sheet_names()) # [u'sheet1', u'sheet2']
    
    sheet1_name = workbook.sheet_names()[2]
    #print(sheet1_name)
    
    sheet1 = workbook.sheet_by_index(1)
    sheet1 = workbook.sheet_by_name(sheet1_name)
    
    #print(sheet1.name, sheet1.nrows, sheet1.ncols)
    begin_flag = 0
    begin_flag_sun = 0
    jb_time = 0
    print("部门","姓名","日期","","时间段","加班时长（小时）","重点工作")
    for i in range(4, sheet1.nrows):
        #print(i)
        name = sheet1.cell_value(i, 7)
        cell_type = sheet1.cell_type(i ,7)
        #print (name, cell_type)

        status = sheet1.cell_value(i,8)
        #if status == "出差":
        #    print("他出差了")
        #if status == "正常":
        #    print("正常上班")

       # print("###",sheet1.cell_value(i,8))
       # print(name[11:13])
       # print(name[14:16])
        hour = int(name[11:13])
        minute = int(name[14:16])
        week = sheet1.cell_value(i, 5)[9:12]
        #检查特殊的加班日
        month = int(name[5:7])
        day = int(name[8:10])
        #print(month, day)
        
        
        #正常加班
        if hour >= 18 and status =="正常":
            if week != "星期日" and week != "星期六":
                jb_time = hour-17 + minute/60
                time_field = "17:00-"+str(hour)+":"+str(minute)
                print("监控应用研发部",sheet1.cell_value(i,0),sheet1.cell_value(i, 5),time_field, round(jb_time,2),"新一代")
        
        #判断周末加班
        if begin_flag == 0:
            if week == "星期六" and status == "出差":
                begin_flag = 1
                #记录开始时间
                week_end_hour = hour
                week_end_minute = minute
               # print(i, "$$$$$$$$$$$$$$$", week, week_end_hour, week_end_minute)
        if begin_flag_sun == 0:
            if week == "星期日" and status == "出差":
                begin_flag_sun = 1
                #记录开始时间
                week_end_hour = hour
                week_end_minute = minute
               # print(i, "$$$$$$$$$$$$$$$", week, week_end_hour, week_end_minute)
        if begin_flag == 1:
            if week == "星期六":
                jb_time = (hour*60+minute - (week_end_hour*60+week_end_minute))/60
               # print(jb_time)
                
        if begin_flag == 1:
            if week != "星期六":
                begin_flag = 0
                #jb_time = (hour*60+minute - (week_end_hour*60+week_end_minute))/60
                print_hour = sheet1.cell_value(i-1, 7)[11:13]
                print_minute = sheet1.cell_value(i-1, 7)[14:16]
                print_field = str(week_end_hour)+":"+str(week_end_minute)+"-"+print_hour+":"+print_minute
                status_new =  sheet1.cell_value(i-1,8)
                print("监控应用研发部",sheet1.cell_value(i-1,0),sheet1.cell_value(i-1, 5), print_field, round(jb_time, 2),"新一代",status_new)

        if begin_flag_sun == 1:
            if week == "星期日":
                jb_time = (hour*60+minute - (week_end_hour*60+week_end_minute))/60
               # print(jb_time)
                
        if begin_flag_sun == 1:
            if week != "星期日":
                begin_flag_sun = 0
                #jb_time = (hour*60+minute - (week_end_hour*60+week_end_minute))/60
                print_hour = sheet1.cell_value(i-1, 7)[11:13]
                print_minute = sheet1.cell_value(i-1, 7)[14:16]
                print_field = str(week_end_hour)+":"+str(week_end_minute)+"-"+print_hour+":"+print_minute
                status_new = sheet1.cell_value(i-1,8)
                print("监控应用研发部",sheet1.cell_value(i-1,0),sheet1.cell_value(i-1, 5),  print_field,  round(jb_time, 2),"新一代", status_new)
               
if __name__ == '__main__':
    read_excel()
