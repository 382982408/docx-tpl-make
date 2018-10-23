# -*- encoding=utf-8 -*-
from openpyxl import load_workbook
from collections import namedtuple
from docxtpl import DocxTemplate
import os


def load_dates(xlsxPath):
	"""
	把xlsx的数据加载到内存中
	xlsxPath: xlsx文件的路径
	"""
	workbook = load_workbook(xlsxPath)
	#booksheet = workbook.active                #获取当前活跃的sheet,默认是第一个sheet
	sheets = workbook.get_sheet_names()         #从名称获取sheet
	booksheet = workbook.get_sheet_by_name(sheets[0])
	 
	rows = booksheet.rows
	titleList = None
	contentList = []
	for row in rows:
	    if not titleList:
	        # 通过第一行的标题生成Record对象
	        line = [col.value for col in row if col.value]
	        titleList = line
	        # 通过namedtuple声称对象
	        # print(titleList)
	        Record = namedtuple('Record', titleList)
	    else:
	        # 先进行检查是否到了终止的地方
	        if row[0].value and row[1].value:
	            # 有数据
	            singleList = [col.value for col in row[:len(titleList)]]
	            record = Record._make(singleList)
	            # print(record)
	            contentList.append(record)
	        else:
	            break;


def makeTplFile(tplPath, namedVlue, outPutdir, isSimple):
	"""
	把内存中的数据做调整，生成对应的word文件
	tplPath: 模板文件中的路径
	namedVlue: 命名的变量在xlsx中的名称
	outPutdir: 输出的路径
	isSimple: 是否是简单的，就是一条生成一个。不是这里改False
	"""
	if not outPutdir.endswith(os.path.sep):
	    outPutdir = outPutdir + os.path.sep
	if not os.path.exists(outPutdir):
	    os.makedirs(outPutdir)
	tpl = DocxTemplate(tplPath)          # 这个是模板的地址
	# 获取模板文件的后缀
	fileType = tplPath[tplPath.rindex('.') + 1]
	# 判断是否是多条一个文件的。
	if isSimple:
	    for record in contentList:
	        context = record._asdict()
	        tpl.render(context)
	        tpl.save(outPutdir + namedVlue + fileType)
	else:
	    from collections import defaultdict
	    complexDict = defaultdict(list)
	    for record in contentList:
	        complexDict[getattr(record, namedVlue)].append(record)
	    for recordList in complexDict.values():
	        context = recordList[0]._asdict()
	        context['items'] = recordList
	        print(context)
	        tpl.render(context)
	        tpl.save(outPutdir + namedVlue + fileType)


if __name__ == '__main__':
	load_dates('/Users/apple/Desktop/任务书数据.xlsx')
	makeTplFile('/Users/apple/Desktop/活动安排-任务书.docx', '参会人名', '/Users/apple/Desktop/docx/', True)