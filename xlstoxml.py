
# pip3 install xlrd

from xml.dom import minidom
from xlrd import open_workbook
import os
import codecs
import sys

languages = ['zh-CN', 'zh-TW', 'zh-HK', 'en', 'es-ES', 'ko-KR', 'pt-PT', 'ja-JP', 'th-TH', 'vi-VN', 'id-ID']

languages_map = {
    'zh-CN': 'values-zh-rCN',
    'zh-TW': 'values-zh-rTW',
    'zh-HK': 'values-zh-rHK',
    'en'   : 'values',
    'es-ES': 'values-es-rES',
    'ko-KR': 'values-ko-rKR',
    'pt-PT': 'values-pt-rPT',
    'ja-JP': 'values-ja-rJP',
    'th-TH': 'values-th-rTH',
    'vi-VN': 'values-vi-rVN',
    'id-ID': 'values-in-rID'
}

def load(path):
    print ('start...')
    (dirpath, tempfilename) = os.path.split(path)
    (app, ext) = os.path.splitext(tempfilename)
    workbook = open_workbook(path)
    doc = minidom.Document()

    mkdir(dirpath + '/' + app)

    # 添加字符串
    for sheet in workbook.sheets():
        for col in range(1, sheet.ncols): 
            # 每一列
            language = sheet.cell(0, col).value.strip()

            if language not in languages:
                print(language + 'not in languages')
                continue

            # 添加xml头
            resources = doc.createElement('resources')
            doc.appendChild(resources)

            # 每一行
            for row_index in range(1 ,sheet.nrows):
                result_placeholder = sheet.cell(row_index, 0).value
                result_content = sheet.cell(row_index, col).value
                # 新建一个文本元素
                text_element = doc.createElement('string')
                text_element.setAttribute('name', result_placeholder)
                text_element.appendChild(doc.createTextNode(formatString(result_content)))
                resources.appendChild(text_element)
            
            xmldir = dirpath + '/' + app + '/' + languages_map[language]
            xmlpath = xmldir + '/strings.xml'
            mkdir(xmldir)
            
            f = codecs.open(xmlpath, 'w', encoding='utf-8')
            f.write(doc.toprettyxml(indent='    '))
            doc.removeChild(resources)

    print('finish ' + dirpath + '/' + app)

def mkdir(path):
    folder = os.path.exists(path) 
    if not folder:                   #判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)            #makedirs 创建文件时如果路径不存在会创建这个路径

def formatString(result_content):
    if "%@" in result_content:
        result_content = result_content.replace('%@','%s')
        print ('替换%@: ' + result_content)

    if "'" in result_content:
        result_content = result_content.replace("'","\\'")
        print ('替换单引号: ' + result_content)
    
    return result_content

load(sys.argv[1])
