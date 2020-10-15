import tempfile
import time
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, LongTable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from io import BytesIO
 
pdfmetrics.registerFont(TTFont('WRYH', 'WeiRuanYaHei.ttf'))  # 默认不支持中文，需要注册字体
pdfmetrics.registerFont(TTFont('WRYHBd', 'WeiRuanYaHeiBD.ttf'))

stylesheet = getSampleStyleSheet()   # 获取样式集
Title = stylesheet['Title']
Title.fontName = 'WRYHBd'
story = []


# 文本基础信息
No='ADBD2020R_20201015114927451857977090'
text_title='项目书1'
author='某某某'

# ~ check_list=['中国学术期刊网络出版总库','中国博士学位论文全文数据库/中国优秀硕士学位论文全文数据库','中国重要会议论文全文数据库','英文数据库(涵盖期刊、博硕、会议的英文数据以及德国Springer、英国Taylor&Francis 期刊数据库等)','图书资源']
# ~ check_list_text=''
# ~ for i in range(0,len(check_list)-1):
	# ~ check_list_text+=check_list[i]+'\n'
# ~ check_list_text+=check_list[-1]

# ~ duration='1900-01-01至'+str(time.strftime("%Y-%m-%d", time.localtime()))

table_data_1 = [['No:'+No,'','','','检测时间：'+str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())),'','',''],
              ['检测文献：',text_title],
              ['项目申请人：',author],
              # ~ ['检测范围：',check_list_text],
              # ~ ['时间范围：',duration]
              ]
table_style_1 = [
    ('SPAN',(0,0),(3,0)),#第一行的合并
    ('SPAN',(4,0),(7,0)),
    ('BACKGROUND',(0, 0),(-1,0),colors.HexColor('#CCFFCC')),  #设置第一行背景颜色
    ('TEXTCOLOR',(0,0),(-1,0),colors.HexColor('#B22222')),#设置第一行字体颜色
    ('ALIGN',(0,0),(-1,0),'LEFT'),  #第一行居左
    ('LINEABOVE',(0,0),(-1,0),1,colors.HexColor('#006400')),#设置第一行上框线
    
    ('SPAN',(1,1),(-1,1)),#后面信息的合并
    ('SPAN',(1,2),(-1,2)),
    # ~ ('SPAN',(1,3),(-1,3)),
    # ~ ('SPAN',(1,4),(-1,4)),
    
    ('TEXTCOLOR',(0,1),(0,-1),colors.HexColor('#006400')),#设置第二行后第一列字体颜色
    ('ALIGN',(0,1),(0,-1),'RIGHT'),  #第二行后第一列居右
    ('TEXTCOLOR',(1,1),(-1,-1),colors.HexColor('#5a5e63')),#设置其他文本字体颜色
    
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # 所有表格上下居中对齐
    ('VALIGN',(0,3),(0,3),'TOP'),  # 检测范围向上对齐
    
    ('FONTNAME', (0, 0), (-1, -1), 'WRYH'),  # 字体
    ('FONTSIZE', (0, 0), (-1, -1), 9),  # 第二行到最后一行的字体大小
]
table1 = Table(data=table_data_1, style=table_style_1, colWidths=60)




#读取复制比数据
f=open("字典数据.txt",'r')
result=eval(f.read())
result_dict={}
for i in range(1,2):
	result_dict['项目书'+str(i+1)]=result


threshold_total=0.5 #设定全文复制比阈值





#复制比概况
table_data_2 = [['检测结果（复制比超过'+"%.f%%"%(threshold_total*100)+'的文章）','','','','','','',''],
              ]

#copied字典记录超过全文检测阈值的文献
copied={}
for k,v in result_dict.items():
	if(v['全文']>threshold_total):
		copied[k]=v['全文']

table_style_2 = [
    ('SPAN',(0,0),(-1,0)),#第一行的合并
    ('BACKGROUND',(0, 0),(-1,-1),colors.HexColor('#CCFFCC')),  #设置全表背景颜色
    ('TEXTCOLOR',(-2,1),(-2,len(copied)),colors.HexColor('#B22222')),#设置复制比的字体颜色
    ('ALIGN',(0,0),(-1,len(copied)),'LEFT'),  #文章表格全部居左
    ('LINEABOVE',(0,0),(-1,0),1,colors.HexColor('#006400')),#设置第一行上框线
    ('LINEBELOW',(0,0),(-1,0),1,colors.HexColor('#006400')),#设置第一行下框线
    
    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),  # 所有表格上下居中对齐
    
    ('FONTNAME',(0,0),(-1,-1),'WRYH'),  # 字体
    ('FONTSIZE',(0,1),(-1,len(copied)),9),  # 文章字体大小
    ('TEXTCOLOR',(0,1),(-3,len(copied)),colors.HexColor('#5a5e63')),#设置文章字体颜色
]

i=1
for k,v in copied.items():
	table_data_2.append([str(i)+'.',k,'','','','',"%.2f%%"%(v*100),''])
	table_style_2.append(('SPAN',(1,i),(5,i)))
	table_style_2.append(('SPAN',(6,i),(7,i)))
	table_style_2.append(('ALIGN',(0,i),(0,i),'RIGHT'))
	i+=1
table_data_2.append([])
table_style_2.append(('LINEABOVE',(0,-1),(-1,-1),0.5,colors.HexColor('#006400')))
table_style_2.append(('BACKGROUND',(0,-1),(-1,-1),colors.white))  #设置空行
table2 = Table(data=table_data_2, style=table_style_2, colWidths=60)
 



#相似文献列表与详情
table_data_3 = [['相似文献列表与详情','','','','','','',''],
              ]

table_style_3 = [
    ('SPAN',(0,0),(-1,0)),#第一行的合并
    ('BACKGROUND',(0, 0),(-1,0),colors.HexColor('#CCFFCC')),  #设置第一行背景颜色
    ('ALIGN',(0,0),(-1,0),'LEFT'),  #第一行居左
    ('LINEABOVE',(0,0),(-1,0),1,colors.HexColor('#006400')),#设置第一行上框线
    ('LINEBELOW',(0,0),(-1,0),1,colors.HexColor('#006400')),#设置第一行下框线
    
    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),  # 所有表格上下居中对齐
    
    ('FONTNAME',(0,0),(-1,-1),'WRYH'),  # 字体
    ('FONTSIZE',(0,1),(-1,-1),9),  # 文章字体大小
    ('TEXTCOLOR',(-1,1),(-1,-1),colors.HexColor('#B22222')),#设置复制比的字体颜色
]


#print(result_dict)

i=1
threshold_para=0.9
for item in copied.keys():
	table_data_3.append([str(i)+'. '+item,'','','','','','文字复制比：',"%.2f%%"%(result_dict[item]['全文']*100)])
	table_style_3.append(('SPAN',(0,len(table_data_3)-1),(5,len(table_data_3)-1)))
	table_style_3.append(('ALIGN',(6,len(table_data_3)-1),(6,len(table_data_3)-1),'RIGHT'))
	table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),10))
	table_style_3.append(('LINEBELOW',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),0.8,colors.black))
	
	
	table_data_3.append(['摘要','','','','','','',"%.2f%%"%(result_dict[item]['摘要']*100)])
	table_style_3.append(('SPAN',(0,len(table_data_3)-1),(6,len(table_data_3)-1)))
	table_style_3.append(('LINEBELOW',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	
	
	
	table_data_3.append(['立项依据与研究内容','','','','','','',"%.2f%%"%(result_dict[item]['立项依据与研究内容']['全文']*100)])
	table_style_3.append(('SPAN',(0,len(table_data_3)-1),(6,len(table_data_3)-1)))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	
	
	table_data_3.append(['','项目的立项依据','','','','','',"%.2f%%"%(result_dict[item]['立项依据与研究内容']['项目的立项依据']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['立项依据与研究内容']['项目的立项依据']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	
	
	table_data_3.append(['','项目的研究内容、研究目标、以及拟解决的关键科学问题','','','','','',"%.2f%%"%(result_dict[item]['立项依据与研究内容']['项目的研究内容、研究目标、以及拟解决的关键科学问题']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['立项依据与研究内容']['项目的研究内容、研究目标、以及拟解决的关键科学问题']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	
	
	table_data_3.append(['','拟采取的研究方案及可行性分析','','','','','',"%.2f%%"%(result_dict[item]['立项依据与研究内容']['拟采取的研究方案及可行性分析']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['立项依据与研究内容']['拟采取的研究方案及可行性分析']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	

	table_data_3.append(['','本项目的特色和创新之处','','','','','',"%.2f%%"%(result_dict[item]['立项依据与研究内容']['本项目的特色和创新之处']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['立项依据与研究内容']['本项目的特色和创新之处']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	

	table_data_3.append(['','年度研究计划及预期研究结果','','','','','',"%.2f%%"%(result_dict[item]['立项依据与研究内容']['年度研究计划及预期研究结果']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['立项依据与研究内容']['年度研究计划及预期研究结果']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	

	table_style_3.append(('LINEBELOW',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	


	table_data_3.append(['研究基础与工作条件','','','','','','',"%.2f%%"%(result_dict[item]['研究基础与工作条件']['全文']*100)])
	table_style_3.append(('SPAN',(0,len(table_data_3)-1),(6,len(table_data_3)-1)))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	
	
	table_data_3.append(['','研究基础','','','','','',"%.2f%%"%(result_dict[item]['研究基础与工作条件']['研究基础']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['研究基础与工作条件']['研究基础']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	

	table_data_3.append(['','工作条件','','','','','',"%.2f%%"%(result_dict[item]['研究基础与工作条件']['工作条件']['全文']*100)])
	table_style_3.append(('SPAN',(1,len(table_data_3)-1),(6,len(table_data_3)-1)))
	for row in result_dict[item]['研究基础与工作条件']['工作条件']['段落']:
		if row[2]>threshold_para:
			table_data_3.append(['','被检文献该章节的第'+str(row[0])+'段与该文献对应章节的第'+str(row[1])+'段','','','','','',"%.2f%%"%(row[2]*100)])
			table_style_3.append(('FONTSIZE',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),8))
			table_style_3.append(('TEXTCOLOR',(0,len(table_data_3)-1),(6,len(table_data_3)-1),colors.HexColor('#5a5e63')))
	table_style_3.append(('LINEBELOW',(1,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	


	table_style_3.append(('LINEBELOW',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),0.4,colors.HexColor('#5a5e63')))	

	table_style_3.append(('LINEBELOW',(0,len(table_data_3)-1),(-1,len(table_data_3)-1),1,colors.black))	

	i+=1


table3 = Table(data=table_data_3, style=table_style_3, colWidths=60)
 
 
story.append(Paragraph("文本复制检测报告单", Title))
story.append(table1)
story.append(table2)
story.append(table3)

 
# file
doc = SimpleDocTemplate('文本复制检测报告单.pdf',topMargin = 50,bottomMargin = 50,title="文本复制检测报告单")
doc.build(story)
