#-*- coding: utf-8 -*-
import xlrd
import xlwt
from datetime import date,datetime
from xlwt import *
workbook = xlrd.open_workbook(r'/data/source/2.xls')
sheet2_name = workbook.sheet_names()[1]
sheet2 = workbook.sheet_by_index(1)
FLAG = True
def read_excel():
	#open file
	#workbook = xlrd.open_workbook(r'/data/source/2.xls')
	#get sheet2
	

	#sheet2 = workbook.sheet_by_index(1)
	rows = sheet2.nrows # 获取行数
	cols = sheet2.ncols
	for x in range(rows-5):
		write_excel(sheet2.row(x+5)[0].value.encode('utf-8'),sheet2.row_values(x+5))
		#print sheet2.row_values(x+5)[1]



	
def write_excel(excel_name,data):
	#被评议人

	wb = xlwt.Workbook()
	ws = wb.add_sheet(u"sheet1",cell_overwrite_ok=True)
	#ws.write(0,22,'',set_borders(2))
	#00000000000000000000000
	row0 = [u'民主评议结果统计工作用表']
	ws.write_merge(0,0,0,3,'',set_style(u'宋体',400,True,1,0))
	ws.write_merge(0,0,3,13,row0,set_style(u'宋体',400,True,1,0))
	#1111111111111111111111111111
	row1 = [u'单位:']
	ws.write_merge(1,1,0,1,row1,set_style(u'宋体',200,False,1,0))
	row1_1 = [u'被评议人:']
	ws.write_merge(1,1,12,12,row1_1,set_style(u'宋体',200,False,1,0))
	row1_1_1 = [u' 年     月     日']
	ws.write(1,14,row1_1_1,set_style(u'宋体',200,False,1,0))
	# 222222222/33333333333333333
	row2 = [u'综合评价']
	ws.write_merge(2,3,0,3,row2,set_style(u'宋体',200,False,22,1,3))
	#ws.write_merge(2,3,0,3,row2,set_pattern(22))
	row2_2 = [u'投票人数',u'有效票数',u'优秀',u'占比',u'称职',u'占比',u'基本称职',u'占比(>=35%为异',u'不称职',u'占比(>=25%为异',u'基本称职＋不称职',u'占比(>=45%为异']
	#投票人数　有效票数　优秀　占比　称职　占比　 基本称职　占比（>=35%为异）　不称职　占比(>=25%为异）基本称职＋不称职　占比(>＝４５ ％为异
	for i in range(0,len(row2_2)):
		ws.write(2,i+4,row2_2[i],set_style(u'宋体',200,False,1,1,3))
	ws.write(2,4,u'投票人数',set_style(u'宋体',200,False,5,1,3))
	ws.write(2,6,u'优秀',set_style(u'宋体',200,False,5,1,3))
	ws.write(2,8,u'称职',set_style(u'宋体',200,False,5,1,3))
	ws.write(2,10,u'基本称职',set_style(u'宋体',200,False,5,1,3))
	ws.write(2,12,u'不称职',set_style(u'宋体',200,False,5,1,3))

	for i in range(12):
		ws.write(3,i+4,'',set_style(u'宋体',200,False,1,1))

	#44444444444444444444444444
	ws.write_merge(4,4,0,15,'',set_style(u'宋体',400,True,1,1))
	#555555555555555555555
	rows5 = [u'测评要素']
	ws.write_merge(5,6,0,3,rows5,set_style(u'宋体',200,False,22,1))
	#ws.write_merge(5,6,0,3,rows5,set_pattern(22))

	rows5_5 = [u'投票人数',u'有效票数',u'优',u'占比',u'良',u'占比',u'中',u'占比',u'差',u'占比(>=20%为异常)',u'中+差',u'占比(>=30%为异常)']
	#for x in range(0,len(rows5_5),2):
	#	ws.write_merge(5,6,x+4,x+4,'x',set_style(u'宋体',200,False))
	for i in range(0,len(rows5_5)):
		#ws.write(5,i+4,rows5_5[i],set_style(u'宋体',200,False))
		ws.write_merge(5,6,i+4,i+4,rows5_5[i],set_style(u'宋体',200,False,1,1))

	ws.write(5,4,u'投票人数',set_style(u'宋体',200,False,5,1))
	ws.write(5,6,u'优',set_style(u'宋体',200,False,5,1))
	ws.write(5,8,u'良',set_style(u'宋体',200,False,5,1))
	ws.write(5,10,u'中',set_style(u'宋体',200,False,5,1))
	ws.write(5,12,u'差',set_style(u'宋体',200,False,5,1))
	#789789787978798
	ws.write_merge(7,7,0,3,u'德',set_style(u'宋体',200,False,1,1))
	ws.write_merge(8,8,0,3,u'能',set_style(u'宋体',200,False,1,1))
	ws.write_merge(9,9,0,3,u'勤',set_style(u'宋体',200,False,1,1))
	ws.write_merge(10,10,0,3,u'绩',set_style(u'宋体',200,False,1,1))
	ws.write_merge(11,11,0,3,u'廉',set_style(u'宋体',200,False,1,1))

	for i in range(5):
		for j in range(13):
			ws.write(i+7,j+3,'',set_style(u'宋体',200,False,1,1))

	#121212121212
	
	ws.write_merge(12,12,0,15,'',set_style(u'宋体',200,False,22,1))
	ws.write_merge(12,12,0,3,'',set_style(u'宋体',200,False,22,1))
	ws.write_merge(12,12,3,13,u'其他文字评价:',set_style(u'宋体',200,False,22,1))
	#13333313333
	ws.write_merge(13,22,0,15,data[45],set_style(u'宋体',200,False,1,1))

	
	#171717
	rows17 = [u'说明:1.该表由工作小组专人使用，每位被评议人一人一表；']
	#\n2.统计时将“投票人数”和个档次得票数填入本表标黄列，由表内公式自动合计“有效票数”并折算占比，异常指数会自动标红；\n3.“其他文字评价”按收集情况原滋原味录入']
	ws.write_merge(23,23,0,15,rows17,set_style(u'宋体',200,False,1,0,3,False))
	rows18 = [u'     2.统计时将“投票人数”和个档次得票数填入本表标黄列，由表内公式自动合计“有效票数”并折算占比，异常指数会自动标红；']
	ws.write_merge(24,24,0,15,rows18,set_style(u'宋体',200,False,1,0,3,False))
	rows19 = [u'     3.“其他文字评价”按收集情况原滋原味录入']
	ws.write_merge(25,25,0,15,rows19,set_style(u'宋体',200,False,1,0,3,False))

	for i in range(6,80):
		ws.write(26,i,"",set_style(u'宋体',200,False,1,0))
		ws.col(i).width = 0x0d00 + i*50
	#ws.write_merge(23,0,15,"xxx",set_style1(u'宋体',200,False,5))

	#ws.write(3,4,5,set_style(u'宋体',200,False,1,0))
	#ws.write(3,5,4,set_style(u'宋体',200,False,1,0))
	#ws.write(3,6,xlwt.Formula("E4-F4"),set_style(u'宋体',200,False,1,0))
	ws.write(3,4,data[37],set_style(u'宋体',200,False,1,1))
	ws.write(3,5,data[37],set_style(u'宋体',200,False,1,1))
	ws.write(3,6,data[38],set_style(u'宋体',200,False,1,1))
	ws.write(3,8,data[39],set_style(u'宋体',200,False,1,1))
	ws.write(3,10,data[40],set_style(u'宋体',200,False,1,1))
	ws.write(3,12,data[41],set_style(u'宋体',200,False,1,1))
	ws.write(3,14,xlwt.Formula("K4+M4"),set_style(u'宋体',200,False,1,1))

	ws.write(7,4,data[2],set_style(u'宋体',200,False,1,1))
	ws.write(7,5,data[2],set_style(u'宋体',200,False,1,1))
	ws.write(7,6,data[3],set_style(u'宋体',200,False,1,1))
	ws.write(7,8,data[4],set_style(u'宋体',200,False,1,1))
	ws.write(7,10,data[5],set_style(u'宋体',200,False,1,1))
	ws.write(7,12,data[6],set_style(u'宋体',200,False,1,1))
	ws.write(7,14,xlwt.Formula("K8+M8"),set_style(u'宋体',200,False,1,1))

	ws.write(8,4,data[9],set_style(u'宋体',200,False,1,1))
	ws.write(8,5,data[9],set_style(u'宋体',200,False,1,1))
	ws.write(8,6,data[10],set_style(u'宋体',200,False,1,1))
	ws.write(8,8,data[11],set_style(u'宋体',200,False,1,1))
	ws.write(8,10,data[12],set_style(u'宋体',200,False,1,1))
	ws.write(8,12,data[13],set_style(u'宋体',200,False,1,1))
	ws.write(8,14,xlwt.Formula("K9+M9"),set_style(u'宋体',200,False,1,1))

	ws.write(9,4,data[16],set_style(u'宋体',200,False,1,1))
	ws.write(9,5,data[16],set_style(u'宋体',200,False,1,1))
	ws.write(9,6,data[17],set_style(u'宋体',200,False,1,1))
	ws.write(9,8,data[18],set_style(u'宋体',200,False,1,1))
	ws.write(9,10,data[19],set_style(u'宋体',200,False,1,1))
	ws.write(9,12,data[20],set_style(u'宋体',200,False,1,1))
	ws.write(9,14,xlwt.Formula("K10+M10"),set_style(u'宋体',200,False,1,1))

	ws.write(10,4,data[23],set_style(u'宋体',200,False,1,1))
	ws.write(10,5,data[23],set_style(u'宋体',200,False,1,1))
	ws.write(10,6,data[24],set_style(u'宋体',200,False,1,1))
	ws.write(10,8,data[25],set_style(u'宋体',200,False,1,1))
	ws.write(10,10,data[26],set_style(u'宋体',200,False,1,1))
	ws.write(10,12,data[27],set_style(u'宋体',200,False,1,1))
	ws.write(10,14,xlwt.Formula("K11+M11"),set_style(u'宋体',200,False,1,1))

	ws.write(11,4,data[30],set_style(u'宋体',200,False,1,1))
	ws.write(11,5,data[30],set_style(u'宋体',200,False,1,1))
	ws.write(11,6,data[31],set_style(u'宋体',200,False,1,1))
	ws.write(11,8,data[32],set_style(u'宋体',200,False,1,1))
	ws.write(11,10,data[33],set_style(u'宋体',200,False,1,1))
	ws.write(11,12,data[34],set_style(u'宋体',200,False,1,1))
	ws.write(11,14,xlwt.Formula("K12+M12"),set_style(u'宋体',200,False,1,1))


	
	
	tmp_data0=data[38]/data[37]*100
	ws.write(3,7,"%.1f%%" %tmp_data0,set_style(u'宋体',200,False,1,1))
	tmp_data0_0=data[39]/data[37]*100
	ws.write(3,9,"%.1f%%" %tmp_data0_0,set_style(u'宋体',200,False,1,1))
	if data[40]/data[37] > 0.35:
		tmp_data=data[40]/data[37]*100
		ws.write(3,11,"%.1f%%" %tmp_data,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data1=data[40]/data[37]*100
		ws.write(3,11,"%.1f%%" %tmp_data1,set_style(u'宋体',200,False,1,1))
	if data[41]/data[37] > 0.25:
		tmp_data2=data[41]/data[37]*100
		ws.write(3,13,"%.1f%%" %tmp_data2,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data3=data[41]/data[37]*100
		ws.write(3,13,"%.1f%%" %tmp_data3,set_style(u'宋体',200,False,1,1))
	if (data[41]+data[40])/data[37] > 0.45: 
		tmp_data4=(data[41]+data[40])/data[37]*100
		ws.write(3,15,"%.1f%%" %tmp_data4,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data5=(data[41]+data[40])/data[37]*100
		ws.write(3,15,"%.1f%%" %tmp_data5,set_style(u'宋体',200,False,1,1))
	tmp_data6=data[3]/data[2]*100
	ws.write(7,7,"%.1f%%" %tmp_data6,set_style(u'宋体',200,False,1,1))
	tmp_data7=data[4]/data[2]*100
	ws.write(7,9,"%.1f%%" %tmp_data7,set_style(u'宋体',200,False,1,1))
	tmp_data8=data[5]/data[2]*100
	ws.write(7,11,"%.1f%%" %tmp_data8,set_style(u'宋体',200,False,1,1))
	if data[6]/data[2] > 0.2:
		tmp_data9=data[6]/data[2]*100
		ws.write(7,13,"%.1f%%" %tmp_data9,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data10=data[6]/data[2]*100
		ws.write(7,13,"%.1f%%" %tmp_data10,set_style(u'宋体',200,False,1,1))
	if (data[5]+data[6])/data[2] > 0.3:
		tmp_data11=(data[5]+data[6])/data[2]*100
		ws.write(7,15,"%.1f%%" %tmp_data11,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data12=(data[5]+data[6])/data[2]*100
		ws.write(7,15,"%.1f%%" %tmp_data12,set_style(u'宋体',200,False,1,1))
	tmp_data13=data[10]/data[9]*100
	ws.write(8,7,"%.1f%%" %tmp_data13,set_style(u'宋体',200,False,1,1))
	tmp_data14=data[11]/data[9]*100
	ws.write(8,9,"%.1f%%" %tmp_data14,set_style(u'宋体',200,False,1,1))
	tmp_data15=data[12]/data[9]*100
	ws.write(8,11,"%.1f%%" %tmp_data15,set_style(u'宋体',200,False,1,1))
	if data[13]/data[9] > 0.2:
		tmp_data16=data[13]/data[9]*100
		ws.write(8,13,"%.1f%%" %tmp_data16,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data17=data[13]/data[9]*100
		ws.write(8,13,"%.1f%%" %tmp_data17,set_style(u'宋体',200,False,1,1))
	if (data[13]+data[12])/data[9] > 0.3:
		tmp_data18=(data[13]+data[12])/data[9]*100
		ws.write(8,15,"%.1f%%" %tmp_data18,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data19=(data[13]+data[12])/data[9]*100
		ws.write(8,15,"%.1f%%" %tmp_data19,set_style(u'宋体',200,False,1,1))
	tmp_data20=data[17]/data[16]*100
	ws.write(9,7,"%.1f%%" %tmp_data20,set_style(u'宋体',200,False,1,1))
	tmp_data21=data[18]/data[16]*100
	ws.write(9,9,"%.1f%%" %tmp_data21,set_style(u'宋体',200,False,1,1))
	tmp_data22=data[19]/data[16]*100
	ws.write(9,11,"%.1f%%" %tmp_data22,set_style(u'宋体',200,False,1,1))
	if data[20]/data[16] > 0.2:
		tmp_data23=data[20]/data[16]*100
		ws.write(9,13,"%.1f%%" %tmp_data23,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data24=data[20]/data[16]*100
		ws.write(9,13,"%.1f%%" %tmp_data24,set_style(u'宋体',200,False,1,1))
	if (data[20]+data[19])/data[16] > 0.3:
		tmp_data25=(data[20]+data[19])/data[16]*100
		ws.write(9,15,"%.1f%%" %tmp_data25,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data26=(data[20]+data[19])/data[16]*100
		ws.write(9,15,"%.1f%%" %tmp_data26,set_style(u'宋体',200,False,1,1))
	tmp_data27=data[24]/data[23]*100
	ws.write(10,7,"%.1f%%" %tmp_data27,set_style(u'宋体',200,False,1,1))
	tmp_data28=data[25]/data[23]*100
	ws.write(10,9,"%.1f%%" %tmp_data28,set_style(u'宋体',200,False,1,1))
	tmp_data29=data[26]/data[23]*100
	ws.write(10,11,"%.1f%%" %tmp_data29,set_style(u'宋体',200,False,1,1))
	if data[27]/data[23] > 0.2:
		tmp_data30=data[27]/data[23]*100
		ws.write(10,13,"%.1f%%" %tmp_data30,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data31=data[27]/data[23]*100
		ws.write(10,13,"%.1f%%" %tmp_data31,set_style(u'宋体',200,False,1,1))
	if (data[27]+data[26])/data[23] > 0.3:
		tmp_data32=(data[27]+data[26])/data[23]*100
		ws.write(10,15,"%.1f%%" %tmp_data32,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data33=(data[27]+data[26])/data[23]*100
		ws.write(10,15,"%.1f%%" %tmp_data33,set_style(u'宋体',200,False,1,1))
	tmp_data34=data[31]/data[30]*100
	ws.write(11,7,"%.1f%%" %tmp_data34,set_style(u'宋体',200,False,1,1))
	tmp_data35=data[32]/data[30]*100
	ws.write(11,9,"%.1f%%" %tmp_data35,set_style(u'宋体',200,False,1,1))
	tmp_data36=data[33]/data[30]*100
	ws.write(11,11,"%.1f%%" %tmp_data36,set_style(u'宋体',200,False,1,1))
	if data[34]/data[30] > 0.2:
		tmp_data37=data[34]/data[30]*100
		ws.write(11,13,"%.1f%%" %tmp_data37,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data38=data[34]/data[30]*100
		ws.write(11,13,"%.1f%%" %tmp_data38,set_style(u'宋体',200,False,1,1))
	if (data[34]+data[33])/data[30] > 0.3:
		tmp_data39=(data[34]+data[33])/data[30]*100
		ws.write(11,15,"%.1f%%" %tmp_data39,set_style(u'宋体',200,False,2,1))
	else:
		tmp_data40=(data[34]+data[33])/data[30]*100
		ws.write(11,15,"%.1f%%" %tmp_data40,set_style(u'宋体',200,False,1,1))


	#clos = sheet2.ncols
	
	#a = []
	#a.append(data)
	#ws.write(8,10,data[2],set_style(u'宋体',200,False,1,0))
	#print data[2]

	#print a[data[2]]
	#for i in int(data[2]):
	#	print i
	
	wb.save("/data/source/"+excel_name+".xls")
def set_style(name,height,bold=False,color=1,border=1,width=3,align=True):
	style = xlwt.XFStyle() #init style

	pattern = Pattern()# color
	pattern.pattern = Pattern.SOLID_PATTERN
	pattern.pattern_fore_colour = color
	
	font = xlwt.Font() #  font
	font.name = name # name 'Times New Roman'
	font.bold = bold # size
	font.color_index = 4
	font.height = height

	borders = Borders()
	borders.right = border#kuang
	borders.top = border
	borders.bottom = border

	style.borders = borders
	style.pattern = pattern
	style.font = font
	if align == True:
		alignment = xlwt.Alignment()
		alignment.horz = xlwt.Alignment.HORZ_CENTER
		alignment.vert = xlwt.Alignment.VERT_CENTER
		style.alignment = alignment



	return style

#def set_style(name,height,bold=False):
#	style = xlwt.XFStyle() #init style
#	font = xlwt.Font() # init font
#	font.name = name # name 'Times New Roman'
#	font.bold = bold # size
#	font.color_index = 4
#	font.height = height 
#	style.font = font
#	return style

#def set_pattern(color):
#	style = xlwt.XFStyle() #init style
#	pattern = Pattern()
#	pattern.pattern = Pattern.SOLID_PATTERN
#	pattern.pattern_fore_colour = color
#	style.pattern = pattern
#	return style

#def set_borders(i):
#	
#	style = xlwt.XFStyle()
#	borders = Borders()
#	borders.right = i
#	borders.top = i
#	borders.bottom = i
#	style.borders = borders
#	return style
def test(name,data):
    print data
if __name__ == '__main__':
	read_excel()
	#write_excel()
