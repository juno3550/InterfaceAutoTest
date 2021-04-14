from action.case_action import execute_case
from util.excel_util import Excel
from util.global_var import *
from util.datetime_util import *
from util.report_util import *
from util.log_util import *
from util.email_util import send_mail


# 执行具体模块的用例sheet（注册sheet、登录sheet等）
def suite_process(excel_file_path, sheet_name):
    # 记录测试结果统计
    global TOTAL_CASE
    global PASS_CASE
    global FAIL_CASE
    # 只要有一条用例失败，则本测试用例集结果算失败
    suite_test_flag = True
    # 第一条接口用例的执行时间则为本测试用例集的执行时间
    first_test_time = False
    # 初始化excel对象
    if isinstance(excel_file_path, Excel):
        excel = excel_file_path
    else:
        excel = Excel(excel_file_path)
    if not excel.get_sheet(sheet_name):
        error("sheet名称【%s】不存在！" % sheet_name)
        return
    # 获取所有行数据
    all_row_datas = excel.get_all_row_data()
    if len(all_row_datas) <= 1:
        error("sheet【】数据不大于1行，停止执行！" % sheet_name)
        return
    # 标题行数据
    head_line = all_row_datas[0]
    # 切换到“测试结果”sheet，以写入测试执行结果
    excel.get_sheet("测试结果明细")
    # 写入标题行
    excel.write_row_data(head_line, None, True, "green")
    # 遍历执行用例
    for data in all_row_datas[1:]:
        # 跳过不需要执行的用例
        if data[MAIN_SHEET_IS_EXECUTE_COL_NO].lower() == "n":
            info("接口用例【%s】无需执行！" % data[CASE_API_NAME_COL_NO])
            continue
        TOTAL_CASE += 1
        # 记录测试时间
        execute_time = get_english_datetime()
        # 用例集测试执行同步模块sheet的第一条用例的执行时间
        if not first_test_time:
            first_test_time = execute_time
        case_result_data = execute_case(data)
        # 标识为n的用例返回None
        if not case_result_data:
            continue
        if case_result_data[CASE_TEST_RESULT_COL_NO] == "fail":
            suite_test_flag = False
        # 获取html测试报告的所需数据
        if case_result_data[CASE_TEST_RESULT_COL_NO] == "pass":
            PASS_CASE += 1
            TEST_RESULT_FOR_REPORT.append([case_result_data[CASE_API_NAME_COL_NO], case_result_data[CASE_REQUEST_DATA_COL_NO],
                                           case_result_data[CASE_RESPONSE_DATA_COL_NO], case_result_data[CASE_TIME_COST_COL_NO],
                                           case_result_data[CASE_ASSERT_KEYWORD_COL_NO], "成功", case_result_data[CASE_EXCEPTION_INFO_COL_NO]])
        else:
            FAIL_CASE += 1
            TEST_RESULT_FOR_REPORT.append([case_result_data[CASE_API_NAME_COL_NO], case_result_data[CASE_REQUEST_DATA_COL_NO],
                                           case_result_data[CASE_RESPONSE_DATA_COL_NO], case_result_data[CASE_TIME_COST_COL_NO],
                                           case_result_data[CASE_ASSERT_KEYWORD_COL_NO], "失败", case_result_data[CASE_EXCEPTION_INFO_COL_NO]])
        # 写入excel测试结果明细
        case_result_data[CASE_TEST_TIME_COL_NO] = execute_time
        excel.write_row_data(case_result_data)
    # 写入excel测试结果统计
    excel.get_sheet("测试结果统计")
    excel.insert_row_data(1, [TOTAL_CASE, PASS_CASE, FAIL_CASE])
    # 返回excel对象，是为了在生成测试报告时再调用save()生成excel版测试结果文件，以此同步和html测试报告相同的时间戳命名
    # suite_false_count 返回本次测试用例集的执行结果
    # first_test_time 作为本测试用例集的执行时间
    return excel, suite_test_flag, first_test_time


# 执行主sheet“测试用例集”
def main_suite_process(excel_file_path, sheet_name):
    # 初始化excel对象
    excel = Excel(excel_file_path)
    if not excel:
        error("excel数据文件【%s】不存在！" % excel_file_path)
        return
    if not excel.get_sheet(sheet_name):
        error("sheet名称【%s】不存在！" % sheet_name)
        return
    # 获取所有行数据
    all_row_datas = excel.get_all_row_data()
    if len(all_row_datas) <= 1:
        error("sheet【%s】数据不大于1行，停止执行！" % sheet_name)
        return
    # 标题行数据
    head_line = all_row_datas[0]
    # 用例步骤行数据
    for data in all_row_datas[1:]:
        # 跳过不需要执行的测试用例集
        if data[MAIN_SHEET_IS_EXECUTE_COL_NO].lower() == "n":
            info("#" * 50 + " 测试用例集【%s】无需执行！ " % data[MAIN_SHEET_SUITE_COL_NO] + "#" * 50 + "\n")
            continue
        if data[MAIN_SHEET_SUITE_COL_NO] not in excel.get_all_sheet():
            error("#" * 50 + " 测试用例集【%s】不存在！ " % data[MAIN_SHEET_SUITE_COL_NO] + "#" * 50 + "\n")
            continue
        info("#" * 50 + " 测试用例集【%s】开始执行 " % data[MAIN_SHEET_SUITE_COL_NO] + "#" * 50)
        excel, suite_test_flag, first_test_time = suite_process(excel, data[MAIN_SHEET_SUITE_COL_NO])
        if suite_test_flag:
            info("#" * 50 + " 测试用例集【%s】执行成功！ " % data[MAIN_SHEET_SUITE_COL_NO] + "#" * 50 + "\n")
            data[MAIN_SHEET_TEST_RESULT_COL_NO] = "pass"
        else:
            info("#" * 50 + " 测试用例集【%s】执行失败！ " % data[MAIN_SHEET_SUITE_COL_NO] + "#" * 50 + "\n")
            data[MAIN_SHEET_TEST_RESULT_COL_NO] = "fail"
        data[MAIN_SHEET_TEST_TIME_COL_NO] = first_test_time
        # 切换到“测试结果明细”sheet，以写入测试执行结果
        excel.get_sheet("测试结果明细")
        # 写入标题行
        excel.write_row_data(head_line, None, True, "red")
        # 写入测试结果
        excel.write_row_data(data)
    # 返回excel对象，是为了在生成测试报告时再调用save()生成excel版测试结果文件，以此同步和html测试报告相同的时间戳命名
    return excel


# 区分模块sheet与主sheet的用例执行，生成测试报告，发送测试报告邮件
def main_process(excel_file_path, sheet_name, excel_report_name, html_report_name, receiver, subject, content):
    if sheet_name == "测试用例集":
        excel_obj = main_suite_process(excel_file_path, sheet_name)
    else:
        excel_obj = suite_process(excel_file_path, sheet_name)[0]
    if not isinstance(excel_obj, Excel):
        error("测试执行失败，已停止生成测试报告！")
    timestamp = get_timestamp()
    xlsx, html = create_xlsx_and_html_report(TEST_RESULT_FOR_REPORT, excel_report_name, html_report_name, excel_obj, timestamp)
    send_mail([xlsx, html], receiver, subject+"_"+timestamp, content)


if __name__ == "__main__":
    # excel_obj = main_suite_process(EXCEL_FILE_PATH, "测试用例集")
    # create_xlsx_and_html_report(TEST_RESULT_FOR_REPORT, "接口测试报告", excel_obj, get_timestamp())
    # suite_process(EXCEL_FILE_PATH, "注册&登录")
    # main_suite_process(EXCEL_FILE_PATH, "测试用例集")
    html_report_name = "接口自动化测试报告"
    receiver = "182230124@qq.com"
    subject = "接口自动化测试报告"
    content = "接口自动化测试报告excel版和html版 请查收附件~"
    main_process(EXCEL_FILE_PATH, "测试用例集", html_report_name, receiver, subject, content)





