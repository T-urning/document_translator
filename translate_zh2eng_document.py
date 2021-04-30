import docx
import json
import time
import yaml
import random
import requests

from hashlib import md5
from tqdm import tqdm

# 加载配置文件
config = yaml.load(open('config.yaml', mode='r', encoding='utf-8'))
# 全局变量
TRANSLATE_API = 'https://fanyi-api.baidu.com/api/trans/vip/translate'
APP_ID = str(config['API']['appid']) 
APP_KEY = config['API']['appkey']

# Generate salt and sign
def make_md5(s, encoding='utf-8'):
    return md5(s.encode(encoding)).hexdigest()


def translate(text, source_lang=None, target_lang=None):
    ''' 
    调用百度翻译 API 将文本 text 从 source_lang 语言翻译到 target_lang 语言，API 网址为：https://fanyi-api.baidu.com/product/113
    '''
    text = text.strip()
    if not text:
        raise ValueError('传入的文本不能为空！')
    if source_lang is None:
        source_lang = 'zh' # 'auto' -> 自动检测语种
    if target_lang is None:
        target_lang = 'en'
    salt = random.randint(32768, 65536)
    sign = make_md5(APP_ID + text + str(salt) + APP_KEY)

    # Build request
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {'appid': APP_ID, 'q': text, 'from': source_lang, 'to': target_lang, 'salt': salt, 'sign': sign}

    # Send request
    r = requests.post(TRANSLATE_API, params=payload, headers=headers)
    result = r.json()
    # 拦截错误
    if 'error_code' in result.keys():
        error_code = result['error_code']
        raise Exception(f'调用百度翻译 API 出现错误，错误码为 {error_code}，' +
            '请根据错误码到 https://fanyi-api.baidu.com/product/113 找相应的解决方法。')

    return result['trans_result'][0]['dst']



if __name__ == '__main__':

    # 待翻译文档路径
    file_path = config['FILE']['file_path'] #'test_document.docx'
    
    # 读取文档
    doc = docx.Document(file_path)
    all_paras = doc.paragraphs  # 文档所有段落（标题+正文+图表题注）
    all_tables = doc.tables # 所有表格

    # 翻译所有段落，包括图片与表格的标题名
    print('处理所有段落，包括图片与表格的标题名...')
    for para in tqdm(all_paras):
        if len(para.runs) < 1:
            continue
        # 记录该段落的字体样式
        style = para.runs[0].style 
        text_zh = para.text.strip()
        # 对段落整体进行翻译
        if text_zh:
            text_en = translate(text_zh)
            para.text = para.text.replace(text_zh, text_en)
        # 设置回原先的样式
        for run in para.runs:
            run.style = style
        # 受翻译API访问频率的限制
        time.sleep(1)
    
    # 翻译表格内容
    print('处理所有表格...')
    for table in tqdm(all_tables): # 所有表格
        for i, row in tqdm(enumerate(table.rows)): # 表格的所有行
            for j, cell in enumerate(row.cells): # 该行的所有元素
                text_zh = cell.text.strip()
                if text_zh:
                    text_en = translate(text_zh)
                    cell.text = cell.text.replace(text_zh, text_en)
                    time.sleep(1)
    # 保存翻译结果              
    save_path = config['FILE']['save_path']             
    doc.save(save_path)
    print('翻译结果已保存。')

