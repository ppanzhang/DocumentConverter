import os
import comtypes.client
from pdf2docx import Converter
import PyPDF2
import re
import pdfplumber


def get_file(input_path, output_path, convertType):
    # 获取所有文件名的列表
    filename_list = os.listdir(input_path)
    # 获取所有 Word 文件名列表
    if convertType == "word2pdf":
        filesNameToBeConverted = [filename for filename in filename_list \
                                  if filename.endswith((".doc", ".docx"))]
    elif convertType == "pdf2word":
        filesNameToBeConverted = [filename for filename in filename_list \
                                  if filename.endswith(".pdf")]

    for inputFileName in filesNameToBeConverted:
        # 分离 Word 文件名称和后缀，转化为 PDF 名称
        if convertType == "word2pdf":
            outputFileName = os.path.splitext(inputFileName)[0] + ".pdf"
        elif convertType == "pdf2word":
            outputFileName = os.path.splitext(inputFileName)[0] + ".docx"
        # 如果当前 Word 文件对应的 PDF 文件存在，则不转化
        if outputFileName in filename_list:
            continue
        # 拼接路径和文件名
        inputFilePath = os.path.join(input_path, inputFileName)
        outFilePath = os.path.join(output_path, outputFileName)
        # 生成器
        yield inputFilePath, outFilePath


def word2pdf(input_path, output_path, convertType):
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 0
    for wordpath, pdfpath in get_file(input_path, output_path, convertType):
        newpdf = word.Documents.Open(wordpath)
        newpdf.SaveAs(pdfpath, FileFormat=17)
        newpdf.Close()


def pdf2word(input_path, output_path, convertType):
    for pdfPath, wordPath in get_file(input_path, output_path, convertType):
        # convert pdf to docx
        cv = Converter(pdfPath)
        cv.convert(wordPath, start=0, end=None)
        cv.close()


def RenamePdfFile(input_path):
    path = input_path.replace('\\', '/')
    paper_name = os.listdir(path)
    for temp in paper_name:
        pdf_file = pdfplumber.open(path + '/' + temp)

        try:
            paper_title = pdf_file.pages[0]
            page_one_text = paper_title.extract_text()

            matched_str = re.search(r'.*\s*.*信', page_one_text).group()
            matched_title = matched_str.split('\n')
            filename_title = matched_title[0].strip(' ')
            # print(filename_title)

            match_EN = re.search(r'2022 \s.*\s', page_one_text).group()
            file_names = match_EN.split('\n')
            file_name = file_names[1]

            match_CN = re.search(r'亲爱的.*,', page_one_text).group()
            match_CN = match_CN.split(",")
            file_name_CN = match_CN[0][3:]

            pdf_file.close()

            new_name = filename_title + '_' + file_name + file_name_CN + '.pdf'
            os.rename(path + '/' + temp, path + '/' + new_name)
        except PyPDF2.utils.PdfReadError:
            pass
        except AttributeError:
            pass


if __name__ == "__main__":
    # 获取当前运行路径
    print("=====================================================")
    print("请选择文档操作类型：")
    print(r"<1> : word to pdf")
    print(r"<2> : pdf to word")
    print(r"<3> ：Rename PDF Files")
    print("=====================================================")
    convertType = input()
    path = os.getcwd()

    if convertType == "1":
        print('请输入需要转换的Word文件目录:')
        input_dir = input()
        print('请输入转换后PDF文件目录:')
        output_dir = input()
        word2pdf(path + '\\' + input_dir, path + '\\' + output_dir, "word2pdf")
    elif convertType == "2":
        print('请输入需要转换的PDF文件目录:')
        input_dir = input()
        print('请输入转换后Word文件目录:')
        output_dir = input()
        pdf2word(path + '\\' + input_dir, path + '\\' + output_dir, "pdf2word")
    elif convertType == "3":
        print('请输入需要重命名PDF文件目录:')
        input_dir = input()
        RenamePdfFile(input_dir)
    else:
        pass
