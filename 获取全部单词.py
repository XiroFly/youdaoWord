import json
import hashlib
from time import sleep
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import requests
import getSync
# 从JSON中提取数据
unique_words = set()
un_word =set()
def extract_data_from_json(json_data):
    items = json_data.get("data", {}).get("items", [])
    for item in items:
        if item.get("d",{})=="en":
            word = item.get("c", "")
            un_word.add((word))
def get_data_from_txt(file_path):
    un_words=set()
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
               # 将每一行解析为 JSON 对象
                data = json.loads(line)
                # 提取 'itemName' 并添加到单词列表中
                if 'itemName' in data:
                    unique_words.add((data['itemName']))
                    un_words.add((data['itemName']))
    return un_words

        
with open(r'C:\Users\fly\Desktop\新建文本文档.txt', 'r', encoding='utf-8') as file:
        json_data = json.load(file)
        extract_data_from_json(json_data)
        un_words1= get_data_from_txt(r"G:\Users\fly\python\爬虫\GMAT_3.txt")
        un_words2=get_data_from_txt(r"G:\Users\fly\python\爬虫\IELTS_3.txt")
        un_words3=get_data_from_txt(r"G:\Users\fly\python\爬虫\SAT_3.txt")
        un_words4=get_data_from_txt(r"G:\Users\fly\python\爬虫\TOEFL_3.txt")
        un_words5=get_data_from_txt(r"G:\Users\fly\python\爬虫\CET4_3.txt")
        un_words6=get_data_from_txt(r"G:\Users\fly\python\爬虫\CET4_MEDIUM.txt")
        with open(r"所有单词.txt",'w',encoding='utf-8') as f:
            i=0
            doc = Document()
            for word in unique_words -un_word-un_words5 -un_words6:
                j=0
                if word in un_words1:
                    j=j+1
                if word in un_words2:
                    j=j+1
                if word in un_words3:
                    j=j+1
                if word in un_words4:
                    j=j+1
                if(j>1):
                    i=i+1
                    res = getSync.get_special1(word)
                    getSync.add_word_to_docx(doc, res["word"], res["usphone"], res["trans"], res["词源"])
            doc.save("youdao_陌生.docx")
            print("\n")
            print(i,end="he")
