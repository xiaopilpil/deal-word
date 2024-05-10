import os
import comtypes.client

# 超参数
# 指定输入文件夹和输出文件夹
input_folder = r'D:\text_word\csv_word\result'
output_folder = r'D:\text_word\csv_word\pdf'


def word_to_pdf(input_file, output_file):
    # 创建 Word 应用程序
    word = comtypes.client.CreateObject('Word.Application')
    # 打开文档
    doc = word.Documents.Open(input_file)
    # 将文档另存为 PDF
    doc.SaveAs(output_file, FileFormat=17)
    # 关闭 Word 应用程序
    word.Quit()


def batch_word_to_pdf(input_folder, output_folder):
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    # 遍历输入文件夹中的每个 Word 文件
    for filename in os.listdir(input_folder):
        if filename.endswith('.docx') or filename.endswith('.doc'):
            input_file = os.path.join(input_folder, filename)
            output_file = os.path.join(output_folder, filename.replace('.docx', '.pdf').replace('.doc', '.pdf'))
            # 转换 Word 文件为 PDF
            word_to_pdf(input_file, output_file)


# 批量将 Word 文件转换为 PDF
batch_word_to_pdf(input_folder, output_folder)
