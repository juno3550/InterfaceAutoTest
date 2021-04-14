from bottle import template
import os
from util.global_var import *
from util.datetime_util import *
from util.log_util import *


# 生成测试报告html模板文件
def report_html(data, html_name, timestamp):
    """
    :param data: 保存测试结果的列表对象
    :param html_name: 报告名称
    """
    template_demo = """
    <!-- CSS goes in the document HEAD or added to your external stylesheet -->
    <style type="text/css">
    table.hovertable {
        font-family: verdana,arial,sans-serif;
        font-size:10px;
        color:#333333;
        border-width: 1px;
        border-color: #999999;
        border-collapse: collapse;
    }
    table.hovertable th {
        background-color:#ff6347;
        border-width: 1px;
        padding: 15px;
        border-style: solid;
        border-color: #a9c6c9;
    }
    table.hovertable tr {
        background-color:#d4e3e5;
    }
    table.hovertable td {
        border-width: 1px;
        padding: 15px;
        border-style: solid;
        border-color: #a9c6c9;
    }
    </style>
    
    <!-- Table goes in the document BODY -->
    
    <head>
    
    <meta http-equiv="content-type" content="txt/html; charset=utf-8" />
    
    </head>
    
    <table class="hovertable">
    <tr>
        <th>接口 URL</th><th>请求数据</th><th>接口响应数据</th><th>接口调用耗时(ms)</th><th>断言词</th><th>测试结果</th><th>异常信息</th>
    </tr>
    % for url,request_data,response_data,test_time,assert_word,result,exception_info in items:
    <tr onmouseover="this.style.backgroundColor='#ffff66';" onmouseout="this.style.backgroundColor='#d4e3e5';">
    
        <td>{{url}}</td><td>{{request_data}}</td><td>{{response_data}}</td><td>{{test_time}}</td><td>{{assert_word}}</td><td>
        % if result == '失败':
        <font color=red>
        % elif result == '成功':
        <font color=green>
        % end
        {{result}}</td>
        <td>{{exception_info}}</td>
    </tr>
    % end
    </table>
    """
    html = template(template_demo, items=data)
    """
    :param template_demo: 渲染的模板名称（可以是字符串对象，也可以是模板文件名）
    :param items: 保存测试结果的列表对象
    :return: 渲染之后的模板（字符串对象）
    """
    # 生成测试报告
    save_dir = os.path.join(TEST_REPORT_SAVE_PATH, get_chinese_date())
    if not os.path.exists(save_dir):
        os.mkdir(save_dir)
    report_name = os.path.join(save_dir, html_name+"_"+timestamp+".html")
    with open(report_name, 'wb') as f:
        f.write(html.encode('utf-8'))
    return os.path.join(save_dir, html_name+"_"+timestamp+".html")


# 将测试结果写入html测试报告
def create_html_report(test_result_data, html_name, timestamp):
    html_name = html_name
    return report_html(test_result_data, html_name, timestamp)


# 将测试结果写入excel数据文件
def create_excel_report(excel_obj, save_name, timestamp):
    return excel_obj.save(save_name, timestamp)


# 同时生成两种测试报告（html和html）
def create_xlsx_and_html_report(test_result_data, excel_name, html_name, excel_obj, timestamp):
    html_report_file_path = create_html_report(test_result_data, html_name, timestamp)
    excel_file_path = create_excel_report(excel_obj, excel_name, timestamp)
    info("生成excel测试报告：{}".format(excel_file_path))
    info("生成html测试报告：{}".format(html_report_file_path))
    return excel_file_path, html_report_file_path