import os.path
import sys
from docx import Document
import larkdoc2docx

def get_save_filename(filename: str) -> str:
    path = os.path.split(filename)[0]
    fname = os.path.basename(filename).replace(".docx", "")
    savefname = "{}_output.docx".format(fname)

    if os.path.isfile(os.path.join(path, savefname)):
        num = 1
        savefname = "{}_output {}.docx".format(fname, num)
        while os.path.isfile(os.path.join(path, savefname)):
            num += 1
            savefname = "{}_output {}.docx".format(fname, num)
    return os.path.join(path, savefname)


if __name__ == '__main__':
    filename = ""
    if len(sys.argv) > 1:
        filename = sys.argv[1]
    if filename is None or not os.path.isfile(filename):
        while not os.path.isfile(filename):
            print("文件不存在，请重新输入")
            filename = input("请输入从飞书导出的word文档路径:")
    save_filename = get_save_filename(filename)

    l2d = larkdoc2docx.larkDoc2Docx()
    l2d.read_template_style('file/template.docx')

    document = Document(filename)
    l2d.add_styles_to_document(document)
    document = l2d.change_larkdoc_style(document)
    document.save(save_filename)
    print("已保存至:{}".format(save_filename))
