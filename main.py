from openpyxl import Workbook
from PyPDF2 import PdfMerger
import pymupdf as pdf
import pytesseract
import pdfplumber
import os


current_path = os.getcwd() #获取当前脚本所在目录
os.chdir("发票")#进入发票目录，以提取相关信息
invoice_num = len(os.listdir())#获取当前发票文件的数量


def rename_pdf():
    """对所有的单个pdf文件按照金额进行重命名"""
    for i in range (invoice_num):
        text = ''
        tem_list1 = []
        tem_list2 = []
        with pdf.open(os.listdir()[i]) as file:
            for j in file:
                tem_text = j.get_text()
                text += tem_text
                tem_text.strip()
                tem_list1.append(tem_text)
        if '¥' in text:
            tem_list1 = text.split('\n')
            for k in range(len(tem_list1)):
                if '¥' in tem_list1[k]:
                    num_str = tem_list1[k].strip("(小写) ").strip().strip("¥")
                    tem_list2.append(float(num_str))

        elif '￥' in text:
            tem_list1 = text.split('\n')
            for k in range(len(tem_list1)):
                if '￥' in tem_list1[k]:
                    num_str = tem_list1[k].strip("(小写) ").strip().strip('￥')
                    tem_list2.append(float(num_str))

        elif text == '':
            print("第",(i+1),"张名称为",os.listdir()[i],"的发票是扫描版pdf,而非电脑生成pdf,此脚本暂时不支持这种模式",sep = '')#表示这不是电脑生成的pdf，而是扫描版的pdf无法提取出内容
            print('该脚本暂时不支持该发票文件，该发票文件也不会计入总pdf文件及总金额之中，请在https://github.com/ZPENG0614/invoice_auto.git中提出issues或是自行修改后提出PR',sep = '')
            tem_list2 = [888888888]
        else:
            print("第",(i+1),"张名称为",os.listdir()[i],"发生未知异常,请在https://github.com/ZPENG0614/invoice_auto.git中提出issues或是自行修改后提出PR")
            tem_list2 = [888888888]


        print(max(tem_list2))
        # print(tem_list1)
        # print(tem_list2)
        if text != '':
            os.rename(os.listdir()[i],str(max(tem_list2))+'.pdf')
            pdfs_num.append(max(tem_list2))
            pdfs_num.sort()
    print(pdfs_num)


def extract_scanned_pdf(order,lang = 'chi_sim'):
    """目前该函数必须依靠外部引擎才能正常工作，无法成为广泛适用的python脚本，待修改......"""
    text = ''
    with pdfplumber.open(os.listdir()[order]) as pdf:
        for page in pdf.pages:
            # 1. 页面转为图片（可加参数：如crop裁剪区域）
            img = page.to_image(resolution=300).original  # resolution=300提升清晰度
            # 2. OCR识别
            text += pytesseract.image_to_string(img, lang=lang) + "\n\n"
    return text







def merge_pdf():
    """合并重命名之后的单个pdf文件"""

    # 创建合并器对象
    merger = PdfMerger()
    # 依次添加所有PDF
    for pdf in pdfs_num:
        merger.append(str(pdf)+'.pdf')
    os.chdir('..')
    # 保存合并后的文件
    merger.write("合并后pdf文件.pdf")
    merger.close()
    print("PDF合并完成！")



def excel_pdf():
    """生成一个excel表格，统计发票的数据"""
    """创建Excel文件并写入数据"""
    # 创建工作簿对象
    wb = Workbook()
    # 获取默认的活动工作表（第一个工作表）
    ws = wb.active
    # 给工作表命名
    ws.title = "发票金额汇总"
    ws['A1'] = '发票金额'
    ws['C3'] = '发票张数'
    ws['C4'] = '发票总金额'
    for i in range(len(pdfs_num)):
        ws['A'+str(i+2)] = str(pdfs_num[i])
    ws['D3'] = str(invoice_num+1)
    ws['D4'] = str(sum(pdfs_num))
    # 保存文件
    wb.save('发票金额数据.xls')


if __name__ == '__main__':
#上一行代码决定了这段程序的运行方式。
#1、直接运行，内置变量__name__将被赋值为__main__,条件成立直接执行
#2、作为模块导入到其他的模块之中，内置变量__name__将被赋值为文件名（不加.py），条件不成立
    pdfs_num = []
    rename_pdf()
    merge_pdf()
    excel_pdf()
















# 豆包算法
# # 检查文本中包含的人民币符号类型
# symbol = '¥' if '¥' in text else '￥' if '￥' in text else None
#
# if symbol:
#     tem_list1 = text.split('\n')
#     for line in tem_list1:  # 直接迭代列表元素，不需要使用索引
#         if symbol in line:
#             # 链式调用strip方法处理字符串
#             amount_str = line.strip("(小写) ").strip().strip(symbol)
#             # 尝试转换为浮点数，防止格式错误导致崩溃
#             try:
#                 tem_list2.append(float(amount_str))
#             except ValueError:
#                 print(f"无法转换为数字: {amount_str}")