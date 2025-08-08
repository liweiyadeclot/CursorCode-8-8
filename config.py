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
PAGE_LOAD_WAIT = 3  # 增加页面加载等待时间
ELEMENT_WAIT = 0.2  # 缩短元素等待时间
BUTTON_CLICK_WAIT = 0.3  # 缩短按钮点击等待时间
RECORD_PROCESS_WAIT = 0.5  # 缩短记录处理等待时间

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
    }
}

# 特殊操作标识
BUTTON_PREFIX = "$"  # 按钮操作的前缀标识
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