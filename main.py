from bussiness_process.main_process import *
from util.global_var import *
from util.report_util import *
from util.datetime_util import *
from util.email_util import send_mail


# 冒烟场景
def smoke_test(report_file_name):
    """
    :param report_file_name: 报告名称与邮件标题
    :return:
    """
    excel_obj, _, _ = suite_process(EXCEL_FILE_PATH, "注册")
    excel_obj, _, _ = suite_process(excel_obj, "登录")
    excel_obj, _, _ = suite_process(excel_obj, "查询博文")
    timestamp = get_timestamp()
    # 生成excel和html的两份测试报告
    excel_file, html_file = create_xlsx_and_html_report(TEST_RESULT_FOR_REPORT, report_file_name,
                                                        report_file_name, excel_obj, timestamp)
    receiver = "182230124@qq.com"  # 收件人
    subject = report_file_name + "_" + timestamp  # 邮件标题
    content = "接口自动化测试报告excel版和html版 请查收附件~"  # 邮件正文
    send_mail([excel_file, html_file], receiver, subject, content)


# 测试用例集测试
def suite_test(report_file_name):
    """
    :param report_file_name: 报告名称与邮件标题
    :return:
    """
    receiver = "182230124@qq.com"  # 收件人
    subject = "接口自动化测试报告_全量测试"  # 邮件标题
    content = "接口自动化测试报告excel版和html版 请查收附件~"  # 邮件正文
    main_process(EXCEL_FILE_PATH, "测试用例集", report_file_name, report_file_name, receiver, subject, content)


if __name__ == "__main__":
    # smoke_test("接口自动化测试报告_冒烟测试")
    suite_test("接口自动化测试报告_全量测试")