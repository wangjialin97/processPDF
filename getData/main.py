import pdfplumber
import pandas as pd
import fitz
import re
import os
from sqlExecute import insertPic,insertXls,insertFileName,searchInDB,searchSubFolder,insertSubFolder
import os.path
from saveWord import get_filename
from judgeAndConvert import convertFile


def find_imag(path, img_path, fileName):
    try:
        checkXO = "/Type(?= */XObject)"    # r"/Type(?=*/XObject)"
        checkIM = "/Subtype(?= */Image)"   # r"/Subtype(?=*/image)"
        pdf = fitz.open(path)
        img_count = 0
        len_XREF=pdf.xref_length()   # lenXREF = doc.xref_length()
        for i in range(1, len_XREF):
            text = pdf.xref_object(i)
            isXObject = re.search(checkXO, text)
            isImage = re.search(checkIM, text)
            if not isXObject or not isImage:
                continue
            img_count += 1
            pix = fitz.Pixmap(pdf, i)
            new_name = fileName + "_img{}.png".format(img_count)
            new_name = new_name.replace(':', '')
            if pix.n < 5:
                pix.writePNG(os.path.join(img_path, new_name))
            else:
                pix0 = fitz.Pixmap(fitz.csRGB, pix)
                pix0.writePNG(os.path.join(img_path, new_name))
        print("提取图片文件完成，共提取了{}张图片".format(img_count))
    except Exception as e:
        print(e)

# 图片入库
def pic2Db(picPath):
    n = 0
    files = os.listdir(picPath)
    for file in files:
        path = os.path.join(picPath, file)
        insertPic(path)
        # print('{}已入库'.format(path))
        n += 1
    print("图片路径入库完成，共入库{}条记录".format(n))

# excel入库
def xls2Db(picPath):
    n = 0
    files = os.listdir(picPath)
    for file in files:
        path = os.path.join(picPath, file)
        insertXls(path)
        print('{}已入库'.format(file))
        n += 1
    print("excel路径入库完成，共入库{}条记录".format(n))

# 保存所有图片
def getPic(fpath,fName, argsSubfolder):
    createDir(fpath + r'/picture')
    # files = os.listdir(fpath)
    # for file in files:
    #     if (file != 'picture') & (file != 'excel'):
    path = os.path.join(fpath, fName)
    filePath = path
    img_path = publicPath + r'/getData/Files/' + argsSubfolder + r'/picture'
    _pdf = fitz.open(filePath)
    _len_XREF = _pdf.xref_length()
    print("目标文件名:{}, 页数:{}, 对象:{}。\n正在提取相关数据，请稍后...".format(fName, len(_pdf), _len_XREF - 1))
    find_imag(filePath, img_path, fName)

def getTable(xlsPath, fName, argsSubfolder):
    createDir(xlsPath + r'/excel')
    # files = os.listdir(xlsPath)
    # for file in files:
    path = os.path.join(xlsPath, fName)
    take_table(path, fName, argsSubfolder)

# 获取pdf表格信息
def take_table(filePath, fileName, argsSubfolder):
    if (fileName != 'picture') & (fileName != 'excel'):
        path = publicPath + r'/getData/Files/' + argsSubfolder + r'/excel'
        try:
            count = 1
            with pdfplumber.open(filePath) as pdf:
                # with pd.ExcelWriter('./pdfXls/{}.xlsx'.format(fileName)) as writer:
                with pd.ExcelWriter(path + '/{}.xlsx'.format(fileName)) as writer:
                    for page in pdf.pages:
                        for table in page.extract_tables():
                            data = pd.DataFrame(table[1:],columns=table[0])
                            data.to_excel(writer,sheet_name=f'sheet{count}')
                            count += 1
                    writer.save()
            print("提取{}表格数据完成".format(fileName))
        except Exception as e:
            print(e)
            print('{}【未能提取成功】'.format(fileName))
            xlsUnsuccessList.append(fileName)

# 创建文件夹
def createDir(path):
    if not os.path.exists(path):
        os.mkdir(path)


# 执行操作Operate
def Op(argsFile,argsSubfile):
    # print('=============开始提取图片=============')
    getPic(publicPath + r'\getData\Files/' + argsSubfile, argsFile, argsSubfile)
    # print('=============图片提取完成=============')
    # print('=============开始提取表格信息=============')
    getTable(publicPath + r'\getData\Files/' + argsSubfile, argsFile, argsSubfile)
    print('未能提取信息表格{}\n'.format(xlsUnsuccessList))
    # print('=============表格信息提取完成=============')

# 数据入库
def InDB(subFolder):
    print('=============图片路径写入数据库=============')
    pic2Db(publicPath + r'/getData/Files/' + subFolder + r'/picture')
    print('=============图片路径写入完成=============')
    print('=============表格路径写入数据库=============')
    xls2Db(publicPath + r'/getData/Files/' + subFolder + r'/excel')
    print('=============表格路径写入完成=============')
    # print('=============文本写入数据库=============')
    # get_filename(publicPath + r'/getData/Files/' + subFolder)
    # print('=============文本写入完成=============')



def getFile(subfolderName):
    n = 0
    files = os.listdir(publicPath + r'\getData\Files/' + subfolderName)
    for file in files:
        result = searchInDB(file)
        if result is None:
            insertFileName(file)
            Op(file, subfolderName)             # 提取picture and excel
            print('=============文本写入数据库=============')
            get_filename(publicPath + r'\getData\Files/' + subfolderName + r'/' + file, file)
            print('=============文本写入完成=============')
            n += 1
            # continue
        else:
            # if file in result:
            print('File"{}"exist'.format(file))
            continue

            # else:
            #     print('6')
            #     insertFileName(file)
            #     Op(file, subfolderName)
            #     n += 1
            #     continue


if __name__ == '__main__':
    publicPath = r'E:\codeRunEnviroment'  # 公共路径
    xlsUnsuccessList = []  # 存储未能提取表格信息的pdf名称
    createDir(publicPath + r'/getData/Files')
    path = publicPath + r'/getData/Files'
    if (os.path.exists(path)):
        # 获取该目录下的所有文件或文件夹目录
        files = os.listdir(path)
        for file in files:
            # 得到该文件下所有目录的路径
            dir = os.path.join(path, file)
            # 判断该路径下是否是文件夹
            if (os.path.isdir(dir)):
                h = os.path.split(dir)
                searchResult = searchSubFolder(h[1])
                if searchResult is None:
                    insertSubFolder(h[1])
                    convertFile(publicPath + r'/getData\Files/' + h[1], h[1])
                    getFile(h[1])
                    InDB(h[1])
                else:
                    if file in searchResult:
                        print('Folder"{}"exist'.format(file))
                        continue
                    else:
                        insertSubFolder(h[1])
                        convertFile(publicPath + r'/getData\Files/' + h[1], h[1])
                        getFile(h[1])

                        continue

    # print("新增处理文件{}个".format(n))
    print("数据成功入库。")