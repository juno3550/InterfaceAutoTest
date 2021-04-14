import os


# 工程根目录
PROJECT_ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 维护唯一数参数的文件路径
UNIQUE_NUM_FILE_PATH = os.path.join(PROJECT_ROOT_DIR, "config", "unique_num.txt")

# 维护接口服务端信息的ini文件路径
INI_FILE_PATH = os.path.join(PROJECT_ROOT_DIR, "config", "server_info.ini")

# excel数据文件
EXCEL_FILE_PATH = os.path.join(PROJECT_ROOT_DIR, "test_data", "interface_test_case.xlsx")

# 维护一个参数化全局变量：供接口关联使用
PARAM_GLOBAL_DICT = {}

# 测试报告保存目录
TEST_REPORT_SAVE_PATH = os.path.join(PROJECT_ROOT_DIR, "test_report")

# excel数据文件用例数据列号
CASE_API_NAME_COL_NO = 1
CASE_SERVER_COL_NO = 2
CASE_METHOD_COL_NO = 3
CASE_REQUEST_DATA_COL_NO = 4
CASE_RESPONSE_DATA_COL_NO = 5
CASE_TIME_COST_COL_NO = 6
CASE_ASSERT_KEYWORD_COL_NO = 7
CASE_RELEVANT_PARAM_COL_NO = 8
CASE_IS_EXECUTE_COL_NO = 9
CASE_TEST_TIME_COL_NO = 10
CASE_TEST_RESULT_COL_NO = 11
CASE_EXCEPTION_INFO_COL_NO = 12

# excel主sheet的用例数据列号
MAIN_SHEET_SUITE_COL_NO = 1
MAIN_SHEET_IS_EXECUTE_COL_NO = 2
MAIN_SHEET_TEST_TIME_COL_NO = 3
MAIN_SHEET_TEST_RESULT_COL_NO = 4

# 存储测试报告需要用的测试结果数据
TEST_RESULT_FOR_REPORT = []

# 日志配置文件路径
LOG_CONF_FILE_PATH = os.path.join(PROJECT_ROOT_DIR, "config", "logger.conf")

# 测试结果统计
TOTAL_CASE = 0
PASS_CASE = 0
FAIL_CASE = 0


if __name__ == "__main__":
    print(PROJECT_ROOT_DIR)
    print(LOG_CONF_FILE_PATH)