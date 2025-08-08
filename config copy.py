# 财务系统自动化处理系统配置文件
# 基于电子科技大学财务综合信息门户

# 财务系统配置
FINANCIAL_SYSTEM_CONFIG = {
    "base_url": "https://cwcx.uestc.edu.cn/WFManager/home.jsp",  # 财务系统基础URL
    "login_url": "https://cwcx.uestc.edu.cn/WFManager/login.jsp",  # 登录页面URL
    "expense_form_url": "https://cwcx.uestc.edu.cn/WFManager/WF_YBX/main2.jsp",  # 智能报销页面URL
    
    # 登录凭据（建议使用环境变量）
    "username": "5130008",
    "password": "Uestc418",
    
    # 页面元素选择器 - 基于实际HTML结构
    "selectors": {
        # 登录页面
        "username_input": "#uid",  # 工号输入框
        "password_input": "#pwd",  # 密码输入框
        "login_button": "#zhLogin",  # 登录按钮
        "captcha_input": "#chkcode1",  # 验证码输入框
        "captcha_image": "img[id='checkcodeImg']",  # 验证码图片
        "captcha_button": "#sjyzm",  # 获取验证码按钮
        
        # 主页元素
        "welcome_message": ".news strong",  # 欢迎信息
        "change_password_button": "#changePwd",  # 修改密码按钮
        "logout_button": "a[href*='logout.jsp']",  # 退出按钮
        
        # 系统导航
        "system_navigator": "#sysNavigator",  # 系统导航区域
        "online_booking": "div[onclick*='WF_YB6']",  # 网上预约
        "online_appointment": "div[onclick*='WF_YB6']",  # 网上预约报账
        "smart_expense": "div[onclick*='WF_YBX']",  # 智能报销
        "new_query": "div[onclick*='WF_CWBS']",  # 新版查询
        "project_auth": "div[onclick*='WF_CA']",  # 项目授权
        "budget": "div[onclick*='WF_CB']",  # 全面预算
        "scholarship": "div[onclick*='WF_GF6_NEW']",  # 奖助学金
        
        # 修改密码弹窗
        "password_dialog": "#divChangePwd",  # 密码修改弹窗
        "new_password1": "#newPwd1",  # 新密码1
        "new_password2": "#newPwd2",  # 新密码2
        "password_strength": "#strength_L, #strength_M, #strength_H",  # 密码强度指示器
        "change_password_submit": "#btChangePwd",  # 确认修改密码
        "cancel_password_change": "#btCancelChange",  # 取消修改密码
        "password_error_msg": "#errMsg",  # 密码错误信息
        
        # 智能报销表单（需要根据实际表单页面调整）
        "expense_form": "form",  # 报销表单
        "project_select": "select[name='project']",  # 项目选择
        "account_select": "select[name='account']",  # 科目选择
        "amount_input": "input[name='amount']",  # 金额输入
        "description_input": "textarea[name='description']",  # 描述输入
        "date_input": "input[name='date']",  # 日期输入
        "category_select": "select[name='category']",  # 类别选择
        "vendor_input": "input[name='vendor']",  # 供应商输入
        "invoice_input": "input[name='invoice']",  # 发票号输入
        "submit_button": "input[type='submit']",  # 提交按钮
        "download_pdf": "a[href*='.pdf']",  # 下载PDF链接
        "success_message": ".success-message",  # 成功消息
        
        # 页面框架
        "main_frame": "#wfMain",  # 主框架
        "sub_system_frame": ".frmSubSystem",  # 子系统框架
        "overlay": "#overlay",  # 遮罩层
        
        # 网上预约报账页面
        "apply_expense_button": "button[btnname='申请报销单'], button[guid='D02B3EF852B84C93B3245737DC749AE4'], button.winBtn.funcButton",  # 申请报销单按钮
        "agree_button": "button[btnname='已阅读并同意'], button.winBtn.funcButton",  # 已阅读并同意按钮
    }
}

# 浏览器配置
BROWSER_CONFIG = {
    "headless": False,  # 是否无头模式运行
    "slow_mo": 1000,    # 操作间隔时间（毫秒）
    "timeout": 30000,   # 超时时间（毫秒）
    "viewport": {
        "width": 1920,
        "height": 1080
    },
    "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

# 文件路径配置
FILE_CONFIG = {
    "download_dir": "downloads",  # PDF下载目录
    "backup_dir": "backups",      # 数据备份目录
    "log_dir": "logs",           # 日志目录
    "expense_data_file": "expenses.json",  # 报销数据文件
    "screenshot_dir": "screenshots"  # 截图目录
}

# 项目配置 - 基于电子科技大学财务系统
PROJECT_CONFIG = {
    "system_name": "电子科技大学财务综合信息门户",
    "default_projects": [
        "科研项目",
        "教学项目", 
        "行政项目",
        "基建项目",
        "其他项目"
    ],
    
    "default_accounts": [
        "差旅费",
        "交通费",
        "餐饮费",
        "办公用品",
        "住宿费",
        "会议费",
        "培训费",
        "设备费",
        "材料费",
        "劳务费",
        "其他费用"
    ],
    
    "default_categories": [
        "住宿费",
        "交通费", 
        "餐饮费",
        "办公用品",
        "会议费",
        "培训费",
        "设备购置",
        "材料费",
        "劳务费",
        "其他"
    ],
    
    # 系统功能模块
    "system_modules": {
        "online_booking": {
            "name": "网上预约",
            "id": "WF_YB6",
            "url": "WF_YB6/main2.jsp"
        },
        "smart_expense": {
            "name": "智能报销", 
            "id": "WF_YBX",
            "url": "WF_YBX/main2.jsp"
        },
        "new_query": {
            "name": "新版查询",
            "id": "WF_CWBS", 
            "url": "WF_CWBS/main2.jsp"
        },
        "project_auth": {
            "name": "项目授权",
            "id": "WF_CA",
            "url": "WF_CA/main2.jsp"
        },
        "budget": {
            "name": "全面预算",
            "id": "WF_CB",
            "url": "WF_CB/main2.jsp"
        },
        "scholarship": {
            "name": "奖助学金",
            "id": "WF_GF6_NEW",
            "url": "WF_GF6_NEW/main2.jsp"
        }
    }
}

# 日志配置
LOGGING_CONFIG = {
    "level": "INFO",
    "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    "file": "financial_system.log"
}

# 自动化配置
AUTOMATION_CONFIG = {
    "retry_count": 3,        # 重试次数
    "retry_delay": 2000,     # 重试延迟（毫秒）
    "wait_timeout": 10000,   # 等待超时（毫秒）
    "screenshot_on_error": True,  # 错误时截图
    "screenshot_dir": "screenshots",  # 截图保存目录
    "page_load_timeout": 30000,  # 页面加载超时
    "navigation_timeout": 30000   # 导航超时
}

# 数据验证配置
VALIDATION_CONFIG = {
    "required_fields": ["project", "account", "amount", "description", "date"],
    "amount_min": 0.01,      # 最小金额
    "amount_max": 1000000,   # 最大金额
    "date_format": "%Y-%m-%d",  # 日期格式
    "username_format": r"^\d{7,8}$",  # 工号格式（7-8位数字）
    "password_min_length": 6  # 密码最小长度
}

# 安全配置
SECURITY_CONFIG = {
    "encrypt_passwords": True,  # 是否加密密码
    "session_timeout": 3600,    # 会话超时时间（秒）
    "max_login_attempts": 3,    # 最大登录尝试次数
    "lockout_duration": 1800    # 锁定时间（秒）
} 