import re
import traceback
import time
import requests
from util.excel_util import Excel
from util.request_util import http_client, get_request_url, get_unique_num, assert_keyword
from util.datetime_util import *
from util.global_var import *
from util.log_util import *


# 对请求数据进行预处理：参数化、函数化
def data_preprocessor(data):
    # 匹配需要调用唯一数函数的参数
    if re.search(r"\$\{unique_\w*\}", data):
        unique_num = get_unique_num()
        # 从用例中获取唯一数的变量名，供后续接口关联使用
        global_num_name = re.search(r"\$\{unique_(\w*)\}", data).group(1)
        # 将调用获取的唯一数的变量名和值存入全局变量中，供后续接口关联使用
        PARAM_GLOBAL_DICT[global_num_name] = unique_num
        data = re.sub(r"\$\{unique_\w*\}", unique_num, data)
    # 匹配需要进行关联的参数
    if re.search(r"\$\{\w+\}", data):
        var = re.search(r"\$\{(\w+)\}", data).group(1)
        data = re.sub(r"\$\{\w+\}", PARAM_GLOBAL_DICT[var], data)
    # 匹配需要进行函数话的参数
    if re.search(r"\$\{\w+?\(.+?\)\}", data):
        func_var = re.search(r"\$\{(\w+?\(.+?\))\}", data).group(1)
        func_result = eval(func_var)
        data = re.sub(r"\$\{(\w+?\(.+?\))\}", func_result, data)
    return data


# 将响应数据需要关联的参数保存进全局变量，供后续接口使用
def data_postprocessor(response_data, revelant_param):
    if not isinstance(revelant_param, str):
        error("数据格式有误！【%s】" % revelant_param)
        error(traceback.format_exc())
    param, regx = revelant_param.split("=")
    # none标识为该条测试数据没有关联数据，因此无需处理
    if regx.lower() == "none":
        return
    if re.search(regx, response_data):
        var_result = re.search(regx, response_data).group(1)
        final_regx = re.sub(r"\(.+\)", var_result, regx)
        info("关联数据【%s】获取成功！" % final_regx)
        PARAM_GLOBAL_DICT[param] = var_result
        return "%s" % final_regx
    else:
        error("关联数据【%s】在响应数据中找不到！" % regx)


# 执行接口用例
def execute_case(data):
    if not isinstance(data, (list, tuple)):
        error("测试用例数据格式有误！测试数据应为列表或元组类型！【%s】" % data)
        data[CASE_EXCEPTION_INFO_COL_NO] = "测试用例数据格式有误！应为列表或元组类型！【%s】" % data
        data[CASE_TEST_RESULT_COL_NO] = "fail"
        return data
    # 该用例无需执行
    if data[CASE_IS_EXECUTE_COL_NO].lower() == "n":
        return
    # 获取请求地址
    url = get_request_url(data[CASE_SERVER_COL_NO])
    # 获取请求接口名称
    api_name = data[CASE_API_NAME_COL_NO]
    info("*" * 40 + " 开始执行接口用例【%s】 " % api_name + "*" * 40)
    # 获取请求方法
    method = data[CASE_METHOD_COL_NO]
    # 替换测试数据
    request_data = data[CASE_REQUEST_DATA_COL_NO]
    info("data before process: %s" % request_data)
    try:
        request_data = data_preprocessor(request_data)
    except:
        error("请求数据【%s】预处理失败！" % request_data)
        error(traceback.format_exc())
        data[CASE_EXCEPTION_INFO_COL_NO] = "请求数据【%s】预处理失败！\n%s" % (request_data, traceback.format_exc())
        data[CASE_TEST_RESULT_COL_NO] = "fail"
        return data
    else:
        info("data after process: %s" % request_data)
        # 数据回写到测试结果
        data[CASE_REQUEST_DATA_COL_NO] = request_data
        # 请求接口并获取响应数据
        start_time = time.time()
        response = http_client(url, api_name, method, request_data)
        api_request_time = time.time() - start_time
        data[CASE_TIME_COST_COL_NO] = int(api_request_time*1000)
        if not isinstance(response, requests.Response):
            error("接口用例【%s】返回的响应对象类型有误！" % api_name)
            data[CASE_TEST_RESULT_COL_NO] = "fail"
            data[CASE_EXCEPTION_INFO_COL_NO] = "接口用例【%s】返回的响应对象类型【%s】有误！" % (api_name, response)
            return data
        else:
            info("接口用例【%s】请求成功！" % api_name)
            data[CASE_API_NAME_COL_NO] = response.url
    try:
        data[CASE_RESPONSE_DATA_COL_NO] = response.text
        info("接口响应数据：{}".format(response.text))
        # 进行断言
        assert_keyword(response, data[CASE_ASSERT_KEYWORD_COL_NO])
        info("接口用例【%s】断言【%s】成功！" % (api_name, data[CASE_ASSERT_KEYWORD_COL_NO]))
        data[CASE_TEST_RESULT_COL_NO] = "pass"
    except:
        error("接口用例【%s】断言【%s】失败！" % (api_name, data[CASE_ASSERT_KEYWORD_COL_NO]))
        error(traceback.format_exc())
        data[CASE_TEST_RESULT_COL_NO] = "fail"
        data[CASE_EXCEPTION_INFO_COL_NO] = traceback.format_exc()
        return data
    try:
        if data[CASE_RELEVANT_PARAM_COL_NO]:
            data[CASE_RELEVANT_PARAM_COL_NO] = data_postprocessor(response.text, data[CASE_RELEVANT_PARAM_COL_NO])
    except:
        error("关联数据【%s】处理失败！" % data[CASE_RELEVANT_PARAM_COL_NO])
        error(traceback.format_exc())
        data[CASE_EXCEPTION_INFO_COL_NO] = "关联数据【%s】处理失败！\n%s" % \
                                           (data[CASE_RELEVANT_PARAM_COL_NO], traceback.format_exc())
        data[CASE_TEST_RESULT_COL_NO] = "fail"
        return data
    data[CASE_TEST_RESULT_COL_NO] = "pass"
    return data


if __name__ == "__main__":
    excel = Excel(EXCEL_FILE_PATH)
    excel.get_sheet("注册")
    datas = excel.get_all_row_data(False)
    for row_data in datas:
        execute_case(row_data)
        # excel.write_row_data(row_data)
        # excel.save()