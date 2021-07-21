import os
import shutil
# import glob
# -*- coding:utf-8 -*-
from win32com.client import Dispatch
# from win32com.client import DipatchEx
from win32com.client import constants
from win32com.client import gencache



def wordtopdf(filelist, targetpath):
    valueList = []
    gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
    # 开始转换
    w = Dispatch("Word.Application")  # 输出Microsoft Word
    (filepath, filename) = os.path.split(filelist)  # 分割文件路径和文件名
    softfilename = os.path.splitext(filename)  # 分割文件名和扩展名
    os.chdir(filepath)  # 切换到新路径：filepath
    doc = os.path.abspath(filename)  # 取filename所在的绝对路径
    os.chdir(targetpath)
    pdfname = softfilename[0] + ".pdf"  # 转换为pdf格式
    output = os.path.abspath(pdfname)
    pdf_name = output
    doc = w.Documents.Open(doc, ReadOnly=1)
    """
        表达式.ExportAsFixedFormat (OutputFileName、ExportFormat、OpenAfterExport、OptimizeFor、Range、From、
        To、Item、IncludeDocProps、KeepIRM、CreateBookmarks、DocStructureTags、BitmapMissingFonts、UseISO19005 1、
        FixedFormatExtClassPtr _)
        其中：
            名称	             必需/可选	    数据类型	                     说明
            OutputFileName	 必需	        String	                     新的 PDF 或 XPS 文件的路径和文件名。
            ExportFormat	 必需	        WdExportFormat	             指定采用 PDF 格式或 XPS 格式。
            Item	         可选	        WdExportItem	             指定导出过程是只包括文本还是包括文本和标记。
            CreateBookmarks	 可选	        WdExportCreateBookmarks	     指定是否导出书签和要导出的书签的类型。
    """
    # 返回”document“对象
    doc.ExportAsFixedFormat(output, constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    if os.path.isfile(pdf_name):
        valueList.append(pdf_name)
    else:
        print('转换失败！')
        return False
    w.Quit(constants.wdDoNotSaveChanges)
    return valueList

# .doc 转 .pdf
def convertFile(filepath, subFile):
    path2 = r'其他文件'
    files = os.listdir(filepath)
    targetpath = r'E:\codeRunEnviroment\getData\Files/' + subFile
    for file in files:
        path = os.path.join(filepath, file)
        boor = path.endswith('.pdf')        # 判断是否为pdf
        boor1 = path.endswith('.docx') or path.endswith('doc')  # 判断是否为doc/docx
        if boor:
            continue
        elif boor1:
            wordtopdf(path, targetpath)     # doc/docx转pdf代码
            os.remove(path)
        else:
            mymovefile(path, path2)    # srcfile 需要移动的文件、dstpath 目的地址


def mymovefile(srcfile, dstpath):  # 移动函数
    if not os.path.isfile(srcfile):
        print("%s not exist!" % (srcfile))
    else:
        fpath, fname = os.path.split(srcfile)  # 分离文件名和路径
        if not os.path.exists(dstpath):
            os.makedirs(dstpath)  # 创建路径
        shutil.move(srcfile, dstpath)  # 移动文件
        print("move %s -> %s" % (srcfile, dstpath + fname))




