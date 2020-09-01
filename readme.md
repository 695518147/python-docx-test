
本文将总结之前在使用python-docx包处理 word 文档时的一些理解与经验。

安装与引入
安装
pip3 install python-docx
引入

# 基本引用 通过 Document 可以创建一个文档对象
from docx import Document
# 子对象引用，docx.shared 中包含诸如"字号"，"颜色"，"行间距"等常用模块
# 非必要，建议只在需要的时候进行引用
from docx.shared import Length, Pt, RGBColor
结构
python-docx将整个文章看做是一个Document对象 官方文档 - Document，其基本结构如下：

每个Document包含许多个代表“段落”的Paragraph对象，存放在document.paragraphs中。
每个Paragraph都有许多个代表"行内元素"的Run对象，存放在paragraph.runs中。


在python-docx中，run是最基本的单位，每个run对象内的文本样式都是一致的，也就是说，在从docx文件生成文档对象时，python-docx会根据样式的变化来将文本切分为一个个的Run对象。

你也可以通过它来处理表格 官方文档 - 表格，基本结构如下：

python-docx将文章中所有的表格都存放在document.tables中
每个Table都有对应的行table. rows、列table. columns和单元格(table. cell())
单元格是最基本的单位，每个单元格又被划分成不同的Paragraph对象，具体内容同上。
表格
常用操作
1. 基本流程
# 导入
from docx import Document
# 从文件创建文档对象
document = Document('./template.docx')
# 显示每段的内容
for p in document.paragraphs:
    print(p.text)
# 添加段落
document.add_paragraph('这是新的段落内容')
# 保存文档
document.save('demo.docx')
2. 搜索并替换

'''
全局内容替换
请确保要替换的内容样式一致

Args:
    doc: 文档对象
    old_text: 要被替换的文本
    new_text: 要替换成的文本
'''
def replace_text(doc, old_text, new_text):
    # 遍历每个段落
    for p in doc.paragraphs:
        # 如果要搜索的内容在该段落
        if old_text in p.text:
            # 使用 runs 替换内容但不改变样式
            # 注意！runs 会根据样式分隔内容，确保被替换内容的样式一致
            for run in p.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)

