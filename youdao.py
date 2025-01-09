import requests
import json
import hashlib
#MD5
def md5_hash(text):
    return hashlib.md5(text.encode('utf-8')).hexdigest()
#制造post里的sign参数
def generate_sign(text):
    v = "webdict"
    _ = "web"
    w = "Mk6hqtUp33DGGtoS63tTJbMUYjRrG1Lu"
    
    # 计算 time
    time = len(text + v) % 10
    
    # 计算 r 和 o
    r = text + v
    o = md5_hash(r)
    
    # 计算 n 和最终的 sign
    n = _ + text + str(time) + w + o
    sign = md5_hash(n)
    
    return sign, time

#cookie转换函数
def parse_cookies(cookie_str):
    cookies = {}
    # 将cookie字符串按分号和空格分割，得到每个cookie对
    cookie_items = cookie_str.split('; ')
    
    for item in cookie_items:
        # 将每个cookie对按照等号分割，得到key和value
        key, value = item.split('=', 1)
        cookies[key] = value
    
    return cookies
url = "https://dict.youdao.com/wordbook/webapi/v2/word/list"
params = {
    "limit": 48,
    "offset": 48,
    "sort": "time",
    "lanTo": "",
    "lanFrom": ""
}
#cookie转换格式
cookie="P_INFO=19853230309|1716722963|1|dict_logon|00&99|null&null&null#shd&370200#10#0|&0|null|19853230309; DICT_PERS=v2|urs-phone-web||DICT||web||-1||1716722964158||61.162.51.121||urs-phoneyd.31036000cd774217b@163.com||gyk4OAOfp40qS0HlGOfqy0JuhfY5RfwFRUW0MJBkfPK064PMw40LJ4RT4Of6ZO4qB0qyOMp4OLJZ0YWOM6FhMw40; DICT_UT=urs-phoneyd.31036000cd774217b@163.com; OUTFOX_SEARCH_USER_ID_NCOO=1983874567.9361298; OUTFOX_SEARCH_USER_ID=1184831676@61.179.137.126; _uetvid=4d4d4800540211efba4cc7557eb057fd; DICT_SESS=v2|R-tfUsHHaJBRMJBOLqy0Tyh4puhLgB0ez0HOMk46uRYmhHeKhMTZ0wFRLkW0Ley0QZ64YWP4pB0YEnHTZhMOM0Y5P4YEnMTF0; DICT_LOGIN=3||1723274421580"
cookies = parse_cookies(cookie)
#http请求头
headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Connection": "keep-alive",
    "Host": "dict.youdao.com",
    "Sec-Ch-Ua": '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-site",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
}
# 请求体
def get_data(text):
    sign, time = generate_sign(text)
    data = {
        "q": text,
        "le": "en",  # 假设目标语言为英文，可以根据需要调整
        "t": time,
        "client": "web",
        "sign": sign,
        "keyfrom": "webdict"
    }
    return data

#拿取单词详细内容
def get_Special(text):
    url="https://dict.youdao.com/jsonapi_s?doctype=json&jsonversion=4"
    
    response1= requests.post(url, headers=headers,cookies=cookies,data=get_data(text))
    if response1.status_code == 200:
    # 解析响应的 JSON 数据
        json_response = response1.json().get("etym",{}).get("etyms",{}).get("zh",[])
        filtered_items=[]
        for item in json_response:
            new_item={
                item.get("word")+":"+item.get("desc") :
                item.get("value")
            }
            filtered_items.append(new_item)
        return(filtered_items)
    else : print("错误\n")
#拿取我的单词列表
def get_list():
    response = requests.get(url, headers=headers, params=params,cookies=cookies)
    if response.status_code == 200:
        data = response.json()
        items = data.get("data", {}).get("itemList", [])
        
        # 只提取所需字段
        filtered_items = []
        for item in items:
            filtered_item = {
                "word": item.get("word"),
                "trans": item.get("trans"),
                "usphone": item.get("usphone"),
                "词源" : get_Special(item.get("word"))
            }
            filtered_items.append(filtered_item)
        return filtered_items

    else:
         print(f"Failed to retrieve data. Status code: {response.status_code}")
#写入到docx
def write_to_docx():
    filtered_items=get_list()
    with open("youdao.txt", "w", encoding="utf-8") as file:
        print("YES")
        json.dump( filtered_items, file, ensure_ascii=False, indent=4)
    
write_to_docx()


