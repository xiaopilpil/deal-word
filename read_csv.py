import csv


def read_csv(filename):
    '''

    :param filename: (str)读取scv的文件名
    :return: (list)返回内容列表
    '''
    line = []
    # 打开 CSV 文件
    with open(filename, newline='', encoding='utf-8-sig') as csvfile:  # 使用 utf-8-sig 编码读取文件，忽略 BOM
        # 创建 CSV 读取器
        csv_reader = csv.reader(csvfile)
        # 遍历 CSV 文件中的每一行
        for i, row in enumerate(csv_reader):
            # 输出每一行数据
            line.append(row)
    return line
