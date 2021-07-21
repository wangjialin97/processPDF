import os.path
# import pyocr
# import importlib
# import sys
# import time

from pdfminer3.pdfparser import PDFParser,PDFDocument
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import PDFPageAggregator
from pdfminer3.layout import LTTextBoxHorizontal, LAParams
from pdfminer3.pdfinterp import PDFTextExtractionNotAllowed
from sqlExecute import insertText,searchInDB


def getTxt(path):
    try:
        text_path = path
        fp = open(text_path, 'rb')
        # 用文件对象创建一个PDF文档分析器
        parser = PDFParser(fp)
        # 创建一个PDF文档
        doc = PDFDocument()
        # 连接分析器，与文档对象
        parser.set_document(doc)
        doc.set_parser(parser)
        # 提供初始化密码，如果没有密码，就创建一个空的字符串
        doc.initialize()
        # 检测文档是否提供txt转换，不提供就忽略
        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed
        else:
            # 创建PDF，资源管理器，来共享资源
            rsrcmgr = PDFResourceManager()
            # 创建一个PDF设备对象
            laparams = LAParams()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            # 创建一个PDF解释其对象
            interpreter = PDFPageInterpreter(rsrcmgr, device)

            # 循环遍历列表，每次处理一个page内容
            # doc.get_pages() 获取page列表
            for page in doc.get_pages():
                interpreter.process_page(page)
                # 接受该页面的LTPage对象
                layout = device.get_result()
                # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
                # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等
                # 想要获取文本就获得对象的text属性，
                for x in layout:
                    if (isinstance(x, LTTextBoxHorizontal)):
                        results = x.get_text()
                        insertText(results)
    except Exception as e:
        print(e)


# 遍历filepath下所有文件
def get_filename(filepath,fileName):
    # files = os.listdir(filepath)
    # for file in files:
    #     if (file != 'picture') & (file != 'excel'):
    #         path = os.path.join(filepath, file)
    #         result = searchInDB(file)
    #         if result is None:
    getTxt(filepath)
    print('{}写入成功'.format(fileName))

