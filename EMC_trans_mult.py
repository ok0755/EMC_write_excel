#coding=utf-8
#author:Zhoubin
#program language:Python2.7
'''
功能：创建EMC.xlsx，汇总EMC报告
使用方法：多份EMC报告与本程序放同一路径，双击执行本程序
          执行前不得打开EMC报告
'''
import docx                     # 导入word读写库
import os
import xlsxwriter               # 导入excel处理库
import multiprocessing          # 导入多进程库
import string


# 遍历当前文件夹docx文件
def get_file():
    current_dir=os.getcwd()
    for Emc_report in os.listdir(current_dir):
        if string.upper(os.path.splitext(Emc_report)[1])=='.DOCX': # 扩展名(大写)='.DOCX'
            yield Emc_report
# 子进程
def multi_func(doc_name):
    current_dir=os.getcwd()
    file_full_path=os.path.join(current_dir,doc_name)
    doc_file = docx.Document(file_full_path)

    try:                                                      # 忽略非EMC报告格式word文档
        table = doc_file.tables[0]                            # word文档第一个表格
        length=len(table.rows)                                # 数据总行数
        ar = [float(table.cell(i,13).text) for i in range(3,length)]  # dB值数列
        index = ar.index(min(ar))
        cell_1 = float(table.cell(index+3,1).text)             # 频率值
        cell_2 = min(ar)                                       # 最小dB值
    except:
        pass
    return doc_name,cell_1,cell_2

# 写入excel表
def write_excel(results):
    xl_book = xlsxwriter.Workbook('EMC.xlsx')                 # 同目录下新建EMC.xlsx
    xl = xl_book.add_worksheet('EMC')                         # 工作表名
    xl.set_column('A:A',25)                                   # 设置excel列宽
    xl.set_column('B:B',15)
    xl.set_column('C:C',15)
    xl.set_column('D:D',15)
    xl.write(0,0,u'文件名')                                   # 写入excel首行
    xl.write(0,1,u'频率')
    xl.write(0,2,u'Margin QP')
    j = 1                                                     # excel行数
    for r in results:
        row_value=r.get()
        xl.write(j,0,row_value[0].decode('gb18030','ignore')) # 繁体中文编码
        xl.write(j,1,row_value[1])
        xl.write(j,2,row_value[2])
        j+=1

    # 单元格条件格式，dB<=0 背景黄色
    format_red_bgcolor = xl_book.add_format({'bg_color':'yellow'})
    xl.conditional_format('C2:C{}'.format(j), {'type': 'cell','criteria': '<=','value': 0, 'format': format_red_bgcolor})
    xl_book.close()                                            # 关闭保存excel
    os.popen('EMC.xlsx')                                       # 打开EMC.xlsx
    os.close

if __name__=='__main__':
    multiprocessing.freeze_support()
    cpu_count = multiprocessing.cpu_count()            # 获取CPU核心数量
    p =  multiprocessing.Pool(processes=cpu_count)     # 进程数=CPU核心数
    results = []
    for doc_name in get_file():
        result = p.apply_async(multi_func,(doc_name,)) # 并行启动多进程
        results.append(result)                         # 缓存结果
    p.close()
    p.join()
    write_excel(results)                                # 调用函数,写入表格