# -*- coding: utf-8 -*-
import sys
import importlib
importlib.reload(sys)

from pdfminer.pdfparser import PDFParser,PDFDocument
from pdfminer.pdfinterp import PDFResourceManager,PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed



def readPDF(path, topath):

	f = open(path, "rb")  #二进制打开
	#创建分析器
	parser = PDFParser(f)
	#创建文档
	pdfFile = PDFDocument()
	#连接分析器与文档
	parser.set_document(pdfFile)
	pdfFile.set_parser(parser)
	
	#提供初始化密码
	pdfFile.initialize()
	
	if not pdfFile.is_extractable:
		raise PDFTextExtractionNotAllowed
	else:
		manager = PDFResourceManager()
		
		laparams = LAParams()
		device = PDFPageAggregator(manager, laparams=laparams)
		interpreter = PDFPageInterpreter(manager, device)
		
		for page in pdfFile.get_pages():
			interpreter.process_page(page)
			layout = device.get_result()
			for x in layout:
				if (isinstance(x, LTTextBoxHorizontal)):
					with open(topath, "a") as f:
						str = x.get_text()
						print(str)
						f.write(str+"\n")
	
	
path = r"E:\lianxi\试验.pdf "

topath = r"E:\lianxi\a.txt "
	
	
readPDF(path, topath)
