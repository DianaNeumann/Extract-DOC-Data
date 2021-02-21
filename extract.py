import sys
import os
import shutil
import win32com
from win32com.client import Dispatch
import docx
import zipfile




# открываем word и сохраняем наш doc файл как docx
def doc_to_docx(path):
    word = win32com.client.Dispatch('word.application')
    word.DisplayAlerts = 0
    word.Visible = False

    doc = word.Documents.Open(path)
    doc.SaveAs(path+'x', 12)
    doc.Close()
    word.Quit()



# извлечение текста 
def extract_text(docx_path):
    docx_path = docx_path
    doc = docx.Document(docx_path)

    path = os.path.abspath(sys.argv[1])
    text_path = path.replace(sys.argv[1],'text.txt')

    fp = open(text_path,'a', encoding='utf-8')
    for p in doc.paragraphs:
        fp.write(p.text+u'\n')
    fp.close()


# извлечение изображений
def extract_image(docx_path, dest_path):
    doc = zipfile.ZipFile(docx_path)
    for info in doc.infolist():
        if info.filename.endswith(('.png', '.jpeg', '.gif')):
            doc.extract(info.filename, dest_path)
            shutil.copy(dest_path+"\\"+info.filename, dest_path+"\\"+ docx_path.split("\\")[-1] + info.filename.split("/")[-1])
    doc.close()



path = os.path.abspath(sys.argv[1])
doc_to_docx(path)
docx_path = path+'x'
img_path = path.replace(sys.argv[1],"")
extract_text(docx_path)
extract_image(docx_path, img_path)

os.remove(docx_path)
