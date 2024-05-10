使用说明：运行csv_word.py文件对word文件识别{{内容}}进行替换。

test.docx格式说明：注意{{内容}}若有下划线必须完全囊括，否则将无法识别到，目前只能识别到表格以及段落的内容。

test.csv格式说明：第一列为文件名，即生成的文件名称，其余列为替换的内容。第一行为被替换的参数，其余行为替换的参数。

运行csv_word.py可用于替换模版的字生成新的word文件。

运行transform_pdf.py可将word文件批处理成pdf文件。

运行前请修改一下这两个.py的参数。

测试环境：python3.8

使用到的库docx、csv、comtypes

```python
pip install -r requirements.txt
```

安装依赖的指令
