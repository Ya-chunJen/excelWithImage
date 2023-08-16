import xlrd
import xlsxwriter
import requests
from io import BytesIO
from urllib.request import urlopen
from PIL import Image
import os
from pathlib import Path

def image_insert(file_path,col_list):
    if not str(file_path).endswith('.xls'):
        raise ValueError('请使用xls文件')
    # 读取数据
    xlsx_read = xlrd.open_workbook(file_path)
    sheet_read = xlsx_read.sheet_by_index(0)
    nrow = sheet_read.nrows
    ncol = sheet_read.ncols

    # 输出文件名
    filename = os.path.basename(file_path)
    out_file_name = "out-" + filename
    out_file_path = os.path.join(os.path.dirname(file_path), out_file_name)

    # 写入数据的excel
    xlsx_write = xlsxwriter.Workbook(out_file_path)
    sheet_write = xlsx_write.add_worksheet("result")

    # col_list： 需要转图片的列，list
    col_list = list(sorted(col_list, key=lambda x:x))    # 需要转图片的列，从小到大排序
    col_max = 0    # 写入时的最大列
    col_begin = 0    # 读取的起始列

    for cl in col_list:
        col_last = cl - 1    # 读取的最后一列
        # 写入不需要转图片的列
        for i in range(0, nrow):
            sheet_write.write_row(i, col_max, sheet_read.row_values(i, col_begin, col_last))
        col_max = col_max + (col_last - col_begin)
        col_begin = col_last + 1
        url_len_max = 0    # 最大的url的个数
        sheet_write.write_row(0, col_max, sheet_read.row_values(0, col_last, cl))
        for i in range(1, nrow):
            sheet_write.set_row(i, 100)    # 设置行高

            # 读取url，生成一个list
            url_str = sheet_read.row_values(i, cl-1, cl)[0]
            url_list = url_str.split("，")
            if url_len_max < len(url_list):
                url_len_max = len(url_list)

            k = 0
            for url in url_list:
                image_data = BytesIO(urlopen(url).read())    # 发送请求获取图片数据

                # 读取图片内容，获得图片大小
                im = Image.open(image_data)
                x_size, y_size = im.size

                # 插入图片
                sheet_write.insert_image(row=i, col=col_max+k, filename=url, options={"image_data": image_data, 'x_scale': 100./x_size, 'y_scale': 100./y_size})
                k += 1
                # os.remove('image.png')
                # break
            # break
        sheet_write.set_column(col_max, col_max+url_len_max, 20)    # 设置图片那几列的列宽
        col_max = col_max + url_len_max

    xlsx_write.close()

if __name__ == '__main__':
    print("这个程序可以将Excel文件中的图片URL，直接转为图片呈现出来。")
    file_path_str = input("请将要处理的文件直接拖进来：")
    file_path_str = file_path_str.strip('\'')
    file_path = Path(file_path_str)
    imgcol_lst_str = input("请输入图片所在列（如有多个，可以以空格分隔）：").strip().split()
    imgcol_lst = [int(num) for num in imgcol_lst_str]
    image_insert(file_path,imgcol_lst)