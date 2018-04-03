from random import randint
import os
from xlutils.copy import copy 
from xlrd import open_workbook 
import xlwt
import xlrd
from datetime import date
'''--------- Inputs--------------'''
def yearmonth():
    return 2018 ,3
def km_chart():
    return '行駛里程&消耗紀錄表107年3月份.xls'
def target_file():
    return '空白車單.xls'
'''-------------------------------'''
def get_road_dis(name):
    road_dic = {'漢民路':3,'大業北路':3,'小港路':4,'平和路':4,'店北路':4,'立群路':4,'學府路':4,'松和路':5,'鳳福路':5}
    return road_dic[name]
def choose_event(eventsNum):
    event_dic ={1:'購買值班人員早餐',
                            2:'購買值班人員早餐，購買值班人員晚餐',
                            3:'購買值班人員早餐，購買值班人員中餐，購買值班人員晚餐'}
    return event_dic[eventsNum]
def choose_time(time):
    time_dic = {1:('07' , str(40+5*randint(0, 3)) ,'08', str(10+5*randint(0, 2))),
                           2:('07' , str(40+5*randint(0, 3)) ,'08', str(10+5*randint(0, 2)),'17', str(20+5*randint(0, 2)),'17', str(45+5*randint(0, 2)) ),
                           3:('07' , str(40+5*randint(0, 3)) ,'08', str(10+5*randint(0, 2)),'11' , str(10+5*randint(0, 2)) ,'11', str(40+5*randint(0, 2)),'17', str(45+5*randint(0, 2)),'18', str(10+5*randint(0, 2)) )}
    return time_dic[time]
def choose_road(km, d):
     road ={ 3:['漢民路','大業北路'],
                     4:['小港路','平和路','店北路','立群路','學府路'],
                     5:['松和路','鳳福路'] }
     weekend = False
     y, m = yearmonth()
     if(date(y, m, d).isoweekday() >= 6):
         weekend  = True
     if(km>=3 and km <=5):
          return (road[km][randint(0, len(road[km])-1)],)
     elif(km==6):
          return road[3][0] , road[3][1]
     elif(km>=7 and km<=9):
          return road[km-4][randint(0, len(road[km-4])-1)], road[4][randint(0, 4)]
     elif(km==10):
          if(weekend):
               return road[5][0] , road[5][1]
          return road[3][0] , road[3][1],  road[4][randint(0, 4)]
     elif(km==11):
          if(randint(0, 1)==1):
               return road[5][randint(0, 1)] , road[3][0] , road[3][1]
          return road[3][randint(0, 1)] , road[4][randint(0, 4)] , road[4][randint(0, 4)]
     return road[3][randint(0, 1)] ,  road[4][randint(0, 4)] , road[5][randint(0, 1)]
def style_initialize():
     #Font
     font=xlwt.Font()
     font.name = '標楷體'
     font.height=0x00F0
     #Alignment
     alignment = xlwt.Alignment()
     alignment.vert = xlwt.Alignment.VERT_CENTER
     alignment.horz = xlwt.Alignment.HORZ_CENTER
     #Borders
     borders = xlwt.Borders()
     borders.top = xlwt.Borders.THIN
     borders.bottom = xlwt.Borders.THIN
     borders.left = xlwt.Borders.THIN
     borders.right = xlwt.Borders.THIN
     #Style
     style=xlwt.XFStyle()
     style.font=font
     style.alignment = alignment
     style.borders = borders
     return style
def style_border(style, top='THIN', bottom='THIN', left='THIN', right='THIN'):
    borders = xlwt.Borders()
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    if top == 'MEDIUM':
        borders.top = xlwt.Borders.MEDIUM
    if bottom == 'MEDIUM':
        borders.bottom = xlwt.Borders.MEDIUM    
    if left == 'MEDIUM':
        borders.left = xlwt.Borders.MEDIUM   
    if right == 'MEDIUM':
        borders.right = xlwt.Borders.MEDIUM
    style.borders = borders
    return style
if __name__ == '__main__':
    xls = xlrd.open_workbook(km_chart())
    file_path = target_file()
    sheetNo = 0
    offset=0
    for ss in range(2):
         print(ss)
         sheet = xls.sheets()[ss]
         sheet_name = sheet.name#表單名稱
         print(sheet.name) #輸出表單名稱(車號)
         ym = sheet.row_values(1)
         year_month = ym[4]
         print(year_month)#年月
         for j in range(4,36):
              values = sheet.row_values(j) #從第幾列開始讀
              if(isinstance(values[1],float)):
                   if(values[0]<10):#check date
                        dates='0'+str(int(values[0]))
                   else:
                        dates=str(int(values[0]))
                   print(sheetNo)
                   print(values)               
                   Roads= choose_road(int(values[3]),int(values[0]))
                   print(Roads)
                   printRoads=''
                   #Write
                   rb = open_workbook(file_path, formatting_info=True)
                   r_sheet = rb.sheet_by_index(sheetNo) # read only copy to introspect the file
                   wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
                   w_sheet = wb.get_sheet(sheetNo) # the sheet to write to within the writable copy
                   style =  style_initialize() #Set style
                   ###
                   dis_temp=0
                   time = choose_time(len(Roads))
                   for i in range(len(Roads)):
                        w_sheet.write(8+i+offset,3,'本隊－'+Roads[i]+'－本隊',style)
                        w_sheet.write(8+i+offset,11,get_road_dis(Roads[i]),style)
                        w_sheet.write(8+i+offset,2,values[1]+dis_temp,style)
                        dis_temp+= get_road_dis(Roads[i])
                        w_sheet.write(8+i+offset,9,values[1]+dis_temp,style)
                        w_sheet.write(8+i+offset,0,time[4*i]+time[4*i+1],style_border(style, left = 'MEDIUM'))
                        w_sheet.write(8+i+offset,8,time[4*i+2]+time[4*i+3],style_border(style))
                        if(i>=len(Roads)-1):
                             printRoads+=Roads[i]
                        else:
                             printRoads+=Roads[i]+'、'                          
                   w_sheet.write(2+offset,2, choose_event(len(Roads)), style)
                   w_sheet.write(3+offset,2, printRoads, style)
                   w_sheet.write(4+offset,2, '自 '+year_month+dates+'日'+time[0]+'時'+time[1]+'分  至  '+year_month+dates+'日'+time[-2]+'時'+time[-1]+'分',style)
                   w_sheet.write(5+offset,2, sheet_name, style)
                   style=style_border(style, bottom = 'MEDIUM')
                   w_sheet.write(11+offset,2,values[3], style)
                   if(offset== 16):
                       offset= 0
                       sheetNo+=1
                   else:
                       offset= 16                   
                   wb.save(file_path)
               
