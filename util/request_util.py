import requests
import json
import hashlib
import traceback
from util.global_var import *
from util.ini_reader import IniParser
from util.log_util import *


# 获取递增的唯一数参数
def get_unique_num():
    with open(UNIQUE_NUM_FILE_PATH, "r+") as f:
        # 先从文件中获取当前的唯一数
        num = f.read()
        # 将唯一数+1再写回文件
        f.seek(0, 0)
        f.write(str(int(num)+1))
    return num


# MD5加密
def md5(string):
    # 创建一个md5 hash对象
    m = hashlib.md5()
    # 对字符串进行md5加密的更新处理，需要指定编码
    m.update(string.encode("utf-8"))
    # 返回十六进制加密结果
    return m.hexdigest()


# 根据接口主机名称映射获取主机的IP和端口
def get_request_url(server_name):
    p = IniParser(INI_FILE_PATH)
    ip = p.get_value(server_name, "ip")
    port = p.get_value(server_name, "port")
    del p
    return "http://%s:%s" % (ip, port)


# 接口请求函数
def http_client(url, api_name, method, data, headers=None, cookies=None):
    # 校验数据是否符合json格式
    try:
        # 字典对象转json字符串
        if isinstance(data, dict):
            data = json.dumps(data)
        elif isinstance(data, str):
            data = json.loads(data)
            data = json.dumps(data)
    except:
        error("接口【%s】json格式有误！" % (url+"/%s/"%api_name))
        traceback.print_exc()
        return traceback.format_exc()
    if method.lower() == "post":
        response = requests.post(url+"/%s/" % api_name, data=data, headers=headers, cookies=cookies)
    elif method.lower() == "get":
        response = requests.get(url+"/%s" % api_name, params=data, headers=headers, cookies=cookies)
    else:
        error("接口【%s】请求方法【%s】有误！" % (url+"/%s/", method))
        return False
    return response


# 断言
def assert_keyword(response, keyword):
    assert keyword in response.text