# 配置文件
import os

# 文件路径配置
EXCEL_FILE = "报销信息.xlsx"
MAPPING_FILE = "标题-ID.xlsx"
DROPDOWN_MAPPING_FILE = "下拉框映射.xlsx"  # 新增：下拉框映射文件
SHEET_NAME = "BaoXiao_sheet"

# 网页配置
TARGET_URL = "https://cwcx.uestc.edu.cn/WFManager/home.jsp"  # 电子科技大学财务系统

# 浏览器配置
HEADLESS = False  # 是否隐藏浏览器窗口
BROWSER_TYPE = "chromium"  # 浏览器类型: chromium, firefox, webkit

# 等待时间配置（秒）
PAGE_LOAD_WAIT = 3  # 页面加载等待时间
ELEMENT_WAIT = 0.2  # 元素等待时间
BUTTON_CLICK_WAIT = 5  # 按钮点击后等待网页加载时间（增加到5秒）
RECORD_PROCESS_WAIT = 0.5  # 记录处理等待时间
SUBJECT_AMOUNT_WAIT = 5  # 科目金额填写前的页面加载等待时间
BANK_CARD_SELECTION_WAIT = 1  # 银行卡选择等待时间（缩减）
BANK_CARD_DIALOG_WAIT = 2  # 银行卡选择弹窗等待时间（缩减）

# 下拉框字段配置（需要根据实际情况调整）
DROPDOWN_FIELDS = {
    "支付方式": {
        "个人转卡": "10",
        "转账汇款": "2", 
        "合同支付": "11",
        "混合支付": "14",
        "冲销其它项目借款": "9",
        "公务卡认证还款": "15"
    },
    "报销类型": {
        "差旅费": "1",
        "办公费": "2",
        "会议费": "3",
        "培训费": "4",
        "设备费": "5"
    },
    "部门": {
        "计算机学院": "CS001",
        "数学学院": "MATH001",
        "物理学院": "PHY001"
    },
    "费用类型": {
        "住宿费": "HOTEL",
        "交通费": "TRANSPORT",
        "餐饮费": "FOOD"
    },
    "审批状态": {
        "待审批": "PENDING",
        "已审批": "APPROVED",
        "已拒绝": "REJECTED"
    },
    "人员类型": {
        "院士": "院士",
        "国家级人才或同等层次人才": "国家级人才或同等层次人才",
        "2级教授": "2级教授",
        "高级职称人员": "高级职称人员",
        "其他人员": "其他人员"
    }
}

# 特殊操作标识
BUTTON_PREFIX = "$"  # 按钮操作的前缀标识
RADIO_BUTTON_PREFIX = "$$"  # radio按钮操作的前缀标识
NAVIGATION_PREFIX = "@"  # 系统导览框操作的前缀标识
CARD_NUMBER_PREFIX = "*"  # 卡号尾号选择前缀

# 子序列处理配置
SUBSEQUENCE_START_COL = "子序列开始"  # 子序列开始列名
SUBSEQUENCE_END_COL = "子序列结束"  # 子序列结束列名

# 日志配置
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
LOG_FILE = "reimbursement_automation.log"

# 子序列相关列名
SUBSEQUENCE_START_COL = "子序列开始"
SUBSEQUENCE_END_COL = "子序列结束"
SEQUENCE_COL = "序号"

# 错误重试配置
MAX_RETRIES = 3
RETRY_DELAY = 1

# 验证配置
VALIDATE_BEFORE_SUBMIT = True  # 提交前是否验证
VALIDATION_TIMEOUT = 10  # 验证超时时间

# 登录相关配置
LOGIN_WAIT_TIME = 5  # 登录后等待时间
CAPTCHA_INPUT_PROMPT = "请输入验证码: "  # 验证码输入提示

# 打印对话框坐标配置
# 这些坐标需要根据实际屏幕分辨率手动获取并填入
PRINT_DIALOG_COORDINATES = {
    "print_button": {
        "x": 1562,  # Chrome打印对话框中保存按钮的X坐标
        "y": 1083   # Chrome打印对话框中保存按钮的Y坐标
    },
    "filepath_input": {
        "x": 720,  # 文件路径输入框的X坐标
        "y": 98   # 文件路径输入框的Y坐标
    },
    "filename_input": {
        "x": 404,  # 文件名输入框的X坐标
        "y": 17   # 文件名输入框的Y坐标
    },
    "save_button": {
        "x": 971,  # 保存按钮的X坐标
        "y": 869   # 保存按钮的Y坐标
    },
    "yes_button": {
        "x": 700,  # "是"按钮的X坐标（可选）
        "y": 450   # "是"按钮的Y坐标（可选）
    }
}

# 打印对话框处理配置
PRINT_OUTPUT_DIR = "pdf_output"  # PDF输出目录
PRINT_DIALOG_WAIT_TIME = 3       # 等待打印对话框出现的时间
SAVE_DIALOG_WAIT_TIME = 2        # 等待保存对话框出现的时间
PRINT_FILE_PATH = r"C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output"  # 打印文件保存路径 