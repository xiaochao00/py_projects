# -*- coding: utf-8 -*-
"""
python3 -m pip install python_docx -i https://pypi.tuna.tsinghua.edu.cn/simple/
python3 -m pip install pypwin32com -i https://pypi.tuna.tsinghua.edu.cn/simple/
"""

import os
import platform
import subprocess

from docx import Document

docPath = os.path.abspath("a.doc")
docxDir = os.path.abspath(os.getcwd())


def docToDocx(doc_path, docx_dir):
    '''将doc转存为docx'''
    assert str(doc_path).endswith("doc")
    if not is_winsystem():
        doc_to_docx_linux(doc_path, docx_dir)
    else:
        doc_to_docx_win(doc_path, docx_dir)


def doc_to_docx_win(doc_path, docx_dir):
    file_name = os.path.splitext(os.path.basename(doc_path))[0]
    docx_path = os.path.join(docx_dir, file_name + ".docx")
    from win32com.client import Dispatch
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(docx_path, FileFormat=12)
    doc.Close()
    word.Quit()
    print("Success transform %s to %s." % (doc_path, docx_path))


def doc_to_docx_linux(doc_path, docx_dir):
    try:
        output = subprocess.check_output(["soffice", "   ", "--invisible", "--convert-to", "docx_parser", doc_path,
                                          "--outdir", docx_dir], stderr=subprocess.STDOUT)
        print(output)
        print("success transform %s to %s.")
        # do something with output
    except subprocess.CalledProcessError as e:
        print(e)
        print("Fail when transform %s to %s.")


def read_docx_dir(docx_dir):
    for docx_path in os.listdir(docx_dir):
        if not docx_path.endswith("docx_parser"):
            continue
        read_docx(docx_path)


def read_docx(docx_path):
    assert str(docx_path).endswith("docx_parser")
    document = Document(docx_path)
    for paragraph in document.paragraphs:
        print(paragraph.text)


def is_winsystem():
    my_system = platform.system()
    if str(my_system).lower().startswith("win"):
        return True
    return False


docToDocx(doc_path=docPath, docx_dir=docxDir)
read_docx_dir(docx_dir=docxDir)
