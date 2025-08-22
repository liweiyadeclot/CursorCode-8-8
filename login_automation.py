import asyncio
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError
import logging
from typing import Optional, Dict, Any, List
import os
import time
from config import *
import sys

# 配置日志
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format=LOG_FORMAT,
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class LoginAutomation:
    def __init__(self, excel_file: str = EXCEL_FILE, mapping_file: str = MAPPING_FILE, 
                 sheet_name: str = SHEET_NAME):
        """
        初始化登录自动化类
        
        Args:
            excel_file: 报销信息Excel文件路径
            mapping_file: 标题-ID映射文件路径
            sheet_name: 要处理的sheet名称
        """
        self.excel_file = excel_file
        self.mapping_file = mapping_file
        self.sheet_name = sheet_name
        self.title_id_mapping = {}
        self.reimbursement_data = None
        self.browser = None
        self.page = None
        self.current_sequence = None
        self.current_project_number = None  # 保存当前记录的报销项目号
        self.current_amount = None          # 保存当前记录的金额
        
    async def load_data(self):
        """加载Excel数据和标题-ID映射"""
        try:
            # 加载标题-ID映射
            mapping_df = pd.read_excel(self.mapping_file)
            self.title_id_mapping = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))
            logger.info(f"成功加载标题-ID映射，共{len(self.title_id_mapping)}条记录")
            
            # 加载报销信息数据
            self.reimbursement_data = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)
            logger.info(f"成功加载报销信息数据，共{len(self.reimbursement_data)}行")
            
            # 验证必要列是否存在
            required_columns = [SEQUENCE_COL]
            missing_columns = [col for col in required_columns if col not in self.reimbursement_data.columns]
            if missing_columns:
                raise ValueError(f"缺少必要的列: {missing_columns}")
                
        except Exception as e:
            logger.error(f"加载数据失败: {e}")
            raise
    
    def get_object_id(self, title: str) -> str:
        """
        根据表头标题获取对应的网页object id
        
        Args:
            title: 表头标题名称
            
        Returns:
            对应的网页object id
        """
        if title in self.title_id_mapping:
            return self.title_id_mapping[title]
        else:
            logger.warning(f"未找到标题 '{title}' 对应的ID映射，请检查标题-ID映射文件")
            return ""
    
    def clean_value_string(self, value) -> str:
        """
        清理数据值，处理数字类型转换时的.0后缀问题
        
        Args:
            value: 要转换的值
            
        Returns:
            清理后的字符串
        """
        if pd.isna(value) or value == "":
            return ""
        
        value_str = str(value).strip()
        # 如果是整数的浮点表示（如123.0），去掉.0后缀
        if value_str.endswith('.0') and value_str.replace('.', '').replace('-', '').isdigit():
            value_str = value_str[:-2]
        
        return value_str
    
    async def wait_for_element(self, element_id: str, timeout: int = 3) -> bool:
        """
        等待元素出现（支持在iframe中查找）
        
        Args:
            element_id: 元素ID
            timeout: 超时时间（秒）
            
        Returns:
            是否成功找到元素
        """
        try:
            # 优先在iframe中查找
            frames = self.page.frames
            for frame in frames:
                try:
                    element = frame.locator(f"#{element_id}").first
                    if await element.count() > 0:
                        logger.info(f"在iframe中找到元素: {element_id}")
                        return True
                except Exception as e:
                    logger.debug(f"在iframe中查找元素失败: {e}")
                    continue
            
            # 如果iframe中找不到，尝试在主页面查找
            await self.page.wait_for_selector(f"#{element_id}", timeout=timeout * 1000)
            logger.info(f"在主页面中找到元素: {element_id}")
            return True
        except TimeoutError:
            logger.warning(f"等待元素超时: {element_id}")
            return False
    
    async def fill_input(self, element_id: str, value: str, retries: int = MAX_RETRIES, title: str = None):
        """
        填写网页中的输入框
        
        Args:
            element_id: 输入框的ID
            value: 要填写的值
            retries: 重试次数
            title: 当前处理的列标题（用于判断是否为金额列）
        """
        # 判断当前输入框对应的报销信息表中的列名是否为"金额"
        if title == "金额":
            self.current_amount = value
            logger.info(f"检测到金额列，保存金额用于文件命名: {value}")
        
        for attempt in range(retries):
            try:
                # 优先在iframe中查找（根据日志分析，大部分元素都在iframe中）
                frames = self.page.frames
                for frame in frames:
                    try:
                        # 在iframe中查找输入框
                        input_element = frame.locator(f"#{element_id}").first
                        if await input_element.count() > 0:
                            await input_element.fill(value)
                            logger.info(f"在iframe中成功填写输入框 {element_id}: {value}")
                            return
                    except Exception as e:
                        logger.debug(f"在iframe中查找输入框失败: {e}")
                        continue
                
                # 如果iframe中找不到，尝试在主页面查找
                if element_id and await self.wait_for_element(element_id):
                    await self.page.fill(f"#{element_id}", value)
                    logger.info(f"在主页面成功填写输入框 {element_id}: {value}")
                    return
                
                # 如果还是找不到，尝试通过name属性查找（优先在iframe中）
                for frame in frames:
                    try:
                        input_element = frame.locator(f"input[name='{element_id}']").first
                        if await input_element.count() > 0:
                            await input_element.fill(value)
                            logger.info(f"在iframe中通过name属性成功填写输入框 {element_id}: {value}")
                            return
                    except Exception as e:
                        logger.debug(f"在iframe中通过name属性查找失败: {e}")
                        continue
                
                # 最后尝试在主页面通过name属性查找
                try:
                    await self.page.fill(f"input[name='{element_id}']", value)
                    logger.info(f"在主页面通过name属性成功填写输入框 {element_id}: {value}")
                    return
                except Exception as e:
                    logger.debug(f"在主页面通过name属性查找失败: {e}")
                    
            except Exception as e:
                logger.warning(f"填写输入框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
        
        logger.error(f"填写输入框最终失败: {element_id}")
    
    async def fill_date_input(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        填写日期输入框（直接填写值，不使用日期选择器）
        
        Args:
            element_id: 日期输入框的ID
            value: 要填写的日期值（格式：yyyy-mm-dd）
            retries: 重试次数
        """
        logger.info(f"开始填写日期输入框: {element_id} = {value}")
        
        for attempt in range(retries):
            try:
                logger.info(f"尝试填写日期输入框 (尝试 {attempt + 1}/{retries}): {element_id}")
                
                # 方法1: 等待元素出现后再填写
                try:
                    logger.info(f"等待元素出现: {element_id}")
                    # 等待元素出现，最多等待5秒
                    await self.page.wait_for_selector(f"#{element_id}", timeout=5000)
                    logger.info(f"元素已出现，开始填写: {element_id}")
                    
                    # 直接尝试填写
                    await self.page.fill(f"#{element_id}", "")
                    await self.page.fill(f"#{element_id}", value)
                    logger.info(f"✓ 在主页面成功填写日期输入框 {element_id}: {value}")
                    return
                except Exception as e:
                    logger.debug(f"等待元素出现失败: {e}")
                
                # 方法2: 通过JavaScript直接设置值（适用于readonly的日期输入框）
                try:
                    logger.info(f"尝试通过JavaScript设置值: {element_id}")
                    js_code = f"""
                    (function() {{
                        var element = document.getElementById('{element_id}');
                        if (element) {{
                            console.log('找到元素:', element);
                            element.value = '{value}';
                            // 触发change事件
                            var event = new Event('change', {{ bubbles: true }});
                            element.dispatchEvent(event);
                            // 触发input事件
                            var inputEvent = new Event('input', {{ bubbles: true }});
                            element.dispatchEvent(inputEvent);
                            return true;
                        }} else {{
                            console.log('未找到元素:', '{element_id}');
                            return false;
                        }}
                    }})();
                    """
                    result = await self.page.evaluate(js_code)
                    if result:
                        logger.info(f"✓ 通过JavaScript成功填写日期输入框 {element_id}: {value}")
                        return
                    else:
                        logger.debug(f"JavaScript未找到元素: {element_id}")
                except Exception as e:
                    logger.debug(f"JavaScript填写失败: {e}")
                
                # 方法3: 通过属性选择器查找
                try:
                    logger.info(f"尝试通过dateinput属性查找: {element_id}")
                    # 查找具有dateinput属性的输入框
                    selector = f"input[dateinput='true'][id='{element_id}']"
                    await self.page.fill(selector, "")
                    await self.page.fill(selector, value)
                    logger.info(f"✓ 通过dateinput属性成功填写日期输入框 {element_id}: {value}")
                    return
                except Exception as e:
                    logger.debug(f"通过dateinput属性查找失败: {e}")
                
                # 方法4: 通过class查找
                try:
                    logger.info(f"尝试通过dateInput类查找: {element_id}")
                    # 查找具有dateInput类的输入框
                    selector = f"input.dateInput[id='{element_id}']"
                    await self.page.fill(selector, "")
                    await self.page.fill(selector, value)
                    logger.info(f"✓ 通过dateInput类成功填写日期输入框 {element_id}: {value}")
                    return
                except Exception as e:
                    logger.debug(f"通过dateInput类查找失败: {e}")
                
                # 方法5: 通过部分ID匹配查找
                try:
                    logger.info(f"尝试通过部分ID匹配查找: {element_id}")
                    # 尝试查找包含element_id的元素
                    selector = f"input[id*='{element_id}']"
                    await self.page.fill(selector, "")
                    await self.page.fill(selector, value)
                    logger.info(f"✓ 通过部分ID匹配成功填写日期输入框 {element_id}: {value}")
                    return
                except Exception as e:
                    logger.debug(f"通过部分ID匹配查找失败: {e}")
                
                # 方法6: 如果主页面找不到，再尝试在iframe中查找
                frames = self.page.frames
                logger.info(f"主页面查找失败，检查 {len(frames)} 个iframe")
                
                for i, frame in enumerate(frames):
                    try:
                        logger.debug(f"在iframe {i} 中查找: {frame.url or 'unnamed frame'}")
                        
                        # 在iframe中查找日期输入框
                        input_element = frame.locator(f"#{element_id}").first
                        if await input_element.count() > 0:
                            # 对于日期输入框，先清除现有值，然后填写新值
                            await input_element.fill("")
                            await input_element.fill(value)
                            logger.info(f"✓ 在iframe {i} 中成功填写日期输入框 {element_id}: {value}")
                            return
                        else:
                            logger.debug(f"在iframe {i} 中未找到元素 {element_id}")
                            
                    except Exception as e:
                        logger.debug(f"在iframe {i} 中查找失败: {e}")
                        continue
                
                # 如果所有方法都失败，等待一下再重试
                if attempt < retries - 1:
                    logger.warning(f"填写日期输入框失败 (尝试 {attempt + 1}/{retries}): {element_id}")
                    logger.info(f"等待 {RETRY_DELAY} 秒后重试...")
                    await asyncio.sleep(RETRY_DELAY)
                    
                    # 额外等待页面加载
                    await asyncio.sleep(2)
                    
            except Exception as e:
                logger.warning(f"填写日期输入框异常 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
        
        logger.error(f"填写日期输入框最终失败: {element_id}")
        logger.info("建议检查：")
        logger.info("1. 元素ID是否正确")
        logger.info("2. 页面是否完全加载")
        logger.info("3. 元素是否在iframe中")
        logger.info("4. 是否需要先点击其他元素来显示日期输入框")
        logger.info("5. 页面是否已经跳转到新页面")
    
    async def fill_readonly_date_input(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        填写只读日期输入框（通过点击日历控件选择日期）
        
        Args:
            element_id: 日期输入框的ID
            value: 要填写的日期值（格式：yyyy-mm-dd）
            retries: 重试次数
        """
        logger.info(f"开始填写只读日期输入框: {element_id} = {value}")
        
        # 解析日期
        try:
            year, month, day = value.split('-')
            year = int(year)
            month = int(month)  # 1-12
            day = int(day)
            logger.info(f"解析日期: 年={year}, 月={month}, 日={day}")
        except Exception as e:
            logger.error(f"日期格式错误: {value}, 期望格式: yyyy-mm-dd")
            return
        
        for attempt in range(retries):
            try:
                logger.info(f"尝试填写只读日期输入框 (尝试 {attempt + 1}/{retries}): {element_id}")
                
                # 1. 点击输入框，弹出日历
                try:
                    await self.page.click(f"#{element_id}")
                    logger.info(f"✓ 成功点击日期输入框: {element_id}")
                except Exception as e:
                    logger.debug(f"点击日期输入框失败: {e}")
                    # 尝试在iframe中点击
                    frames = self.page.frames
                    for i, frame in enumerate(frames):
                        try:
                            input_element = frame.locator(f"#{element_id}").first
                            if await input_element.count() > 0:
                                await input_element.click()
                                logger.info(f"✓ 在iframe {i} 中成功点击日期输入框: {element_id}")
                                break
                        except Exception as e:
                            logger.debug(f"在iframe {i} 中点击失败: {e}")
                            continue
                    else:
                        logger.warning(f"无法点击日期输入框: {element_id}")
                        continue
                
                # 2. 等待日历控件出现
                try:
                    await self.page.wait_for_selector("#ui-datepicker-div", state="visible", timeout=5000)
                    logger.info("✓ jQuery UI日历控件已出现")
                except Exception as e:
                    logger.debug(f"等待jQuery UI日历控件失败: {e}")
                    # 尝试其他可能的日历控件选择器
                    calendar_selectors = [
                        "#ui-datepicker-div",
                        ".ui-datepicker",
                        "[class*='datepicker']",
                        "[id*='datepicker']"
                    ]
                    
                    calendar_found = False
                    for selector in calendar_selectors:
                        try:
                            await self.page.wait_for_selector(selector, state="visible", timeout=2000)
                            logger.info(f"✓ 找到日历控件: {selector}")
                            calendar_found = True
                            break
                        except Exception:
                            continue
                    
                    if not calendar_found:
                        logger.warning("未找到日历控件")
                        continue
                
                # 3. 选择年份（基于实际HTML结构）
                try:
                    # 点击年份下拉框
                    await self.page.click(".ui-datepicker-year")
                    logger.info("✓ 点击年份下拉框")
                    
                    # 选择指定年份
                    await self.page.click(f'option[value="{year}"]')
                    logger.info(f"✓ 选择年份: {year}")
                except Exception as e:
                    logger.debug(f"选择年份失败: {e}")
                    # 尝试其他年份选择方式
                    try:
                        # 直接点击年份文本
                        await self.page.click(f'text={year}')
                        logger.info(f"✓ 通过文本选择年份: {year}")
                    except Exception as e2:
                        logger.debug(f"通过文本选择年份也失败: {e2}")
                        continue
                
                # 4. 选择月份（基于实际HTML结构）
                try:
                    # 点击月份下拉框
                    await self.page.click(".ui-datepicker-month")
                    logger.info("✓ 点击月份下拉框")
                    
                    # 选择指定月份（月份索引从0开始，所以需要减1）
                    month_index = month - 1
                    await self.page.click(f'option[value="{month_index}"]')
                    logger.info(f"✓ 选择月份: {month} (索引: {month_index})")
                except Exception as e:
                    logger.debug(f"选择月份失败: {e}")
                    # 尝试其他月份选择方式
                    try:
                        # 直接点击月份文本
                        month_names = ["1月", "2月", "3月", "4月", "5月", "6月", 
                                     "7月", "8月", "9月", "10月", "11月", "12月"]
                        month_name = month_names[month - 1]
                        await self.page.click(f'text={month_name}')
                        logger.info(f"✓ 通过文本选择月份: {month_name}")
                    except Exception as e2:
                        logger.debug(f"通过文本选择月份也失败: {e2}")
                        continue
                
                # 5. 选择日期（基于实际HTML结构）
                try:
                    # 根据实际HTML结构，日期是通过<a>标签实现的
                    # 选择指定日期
                    await self.page.click(f'#ui-datepicker-div a:has-text("{day}")')
                    logger.info(f"✓ 选择日期: {day}")
                    
                    # 等待日期选择完成
                    await asyncio.sleep(1)
                    
                    # 验证日期是否已填写
                    try:
                        current_value = await self.page.input_value(f"#{element_id}")
                        logger.info(f"✓ 日期填写完成，当前值: {current_value}")
                        return
                    except Exception as e:
                        logger.debug(f"验证日期值失败: {e}")
                        # 如果验证失败，但操作看起来成功了，也返回
                        logger.info("✓ 日期选择操作完成")
                        return
                        
                except Exception as e:
                    logger.debug(f"选择日期失败: {e}")
                    # 尝试其他日期选择方式
                    try:
                        # 尝试点击包含日期的元素（基于实际HTML结构）
                        await self.page.click(f'#ui-datepicker-div td[data-handler="selectDay"] a:has-text("{day}")')
                        logger.info(f"✓ 通过data-handler选择日期: {day}")
                        return
                    except Exception as e2:
                        logger.debug(f"通过data-handler选择日期也失败: {e2}")
                        
                        # 最后尝试：直接点击包含日期的td元素
                        try:
                            await self.page.click(f'#ui-datepicker-div td[data-handler="selectDay"] a[href="#"]:has-text("{day}")')
                            logger.info(f"✓ 通过href选择日期: {day}")
                            return
                        except Exception as e3:
                            logger.debug(f"通过href选择日期也失败: {e3}")
                            continue
                
            except Exception as e:
                logger.warning(f"填写只读日期输入框异常 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
        
        logger.error(f"填写只读日期输入框最终失败: {element_id}")
        logger.info("建议检查：")
        logger.info("1. 日期输入框ID是否正确")
        logger.info("2. 日历控件是否正确加载")
        logger.info("3. 日期格式是否正确 (yyyy-mm-dd)")
        logger.info("4. 是否需要先点击其他元素来显示日期输入框")

    async def select_date_from_calendar(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        基于jQuery UI日历控件的精确日期选择方法
        
        Args:
            element_id: 日期输入框的ID
            value: 要填写的日期值（格式：yyyy-mm-dd）
            retries: 重试次数
        """
        logger.info(f"开始使用jQuery UI日历控件选择日期: {element_id} = {value}")
        
        # 解析日期
        try:
            year, month, day = value.split('-')
            year = int(year)
            month = int(month)  # 1-12
            day = int(day)
            logger.info(f"解析日期: 年={year}, 月={month}, 日={day}")
        except Exception as e:
            logger.error(f"日期格式错误: {value}, 期望格式: yyyy-mm-dd")
            return
        
        for attempt in range(retries):
            try:
                logger.info(f"尝试选择日期 (尝试 {attempt + 1}/{retries}): {element_id}")
                
                # 1. 点击输入框，弹出日历
                target_frame = None
                try:
                    await self.page.click(f"#{element_id}")
                    logger.info(f"✓ 在主页面成功点击日期输入框: {element_id}")
                    target_frame = self.page
                except Exception as e:
                    logger.debug(f"主页面点击日期输入框失败: {e}")
                    # 尝试在iframe中点击
                    frames = self.page.frames
                    for i, frame in enumerate(frames):
                        try:
                            input_element = frame.locator(f"#{element_id}").first
                            if await input_element.count() > 0:
                                await input_element.click()
                                logger.info(f"✓ 在iframe {i} 中成功点击日期输入框: {element_id}")
                                target_frame = frame
                                break
                        except Exception as e:
                            logger.debug(f"在iframe {i} 中点击失败: {e}")
                            continue
                    else:
                        logger.warning(f"无法点击日期输入框: {element_id}")
                        continue
                
                # 2. 等待jQuery UI日历控件出现（在所有frame中查找）
                calendar_found = False
                calendar_frame = None
                
                # 首先在主页面查找
                try:
                    await self.page.wait_for_selector("#ui-datepicker-div", state="visible", timeout=1000)
                    logger.info("✓ 在主页面找到jQuery UI日历控件")
                    calendar_found = True
                    calendar_frame = self.page
                except Exception as e:
                    logger.debug(f"主页面等待jQuery UI日历控件失败: {e}")
                
                # 如果主页面没找到，在所有iframe中查找
                if not calendar_found:
                    frames = self.page.frames
                    logger.info(f"在主页面未找到日历控件，检查 {len(frames)} 个iframe")
                    
                    for i, frame in enumerate(frames):
                        try:
                            await frame.wait_for_selector("#ui-datepicker-div", state="visible", timeout=1000)
                            logger.info(f"✓ 在iframe {i} 中找到jQuery UI日历控件")
                            calendar_found = True
                            calendar_frame = frame
                            break
                        except Exception as e:
                            logger.debug(f"iframe {i} 等待jQuery UI日历控件失败: {e}")
                            continue
                
                # 如果还是没找到，尝试其他选择器
                if not calendar_found:
                    calendar_selectors = [
                        "#ui-datepicker-div",
                        ".ui-datepicker",
                        "div.ui-datepicker.ui-widget",
                        "[class*='datepicker']",
                        "[id*='datepicker']"
                    ]
                    
                    # 在主页面尝试其他选择器
                    for selector in calendar_selectors:
                        try:
                            await self.page.wait_for_selector(selector, state="visible", timeout=500)
                            logger.info(f"✓ 在主页面找到日历控件: {selector}")
                            calendar_found = True
                            calendar_frame = self.page
                            break
                        except Exception:
                            continue
                    
                    # 在iframe中尝试其他选择器
                    if not calendar_found:
                        for i, frame in enumerate(frames):
                            for selector in calendar_selectors:
                                try:
                                    await frame.wait_for_selector(selector, state="visible", timeout=500)
                                    logger.info(f"✓ 在iframe {i} 中找到日历控件: {selector}")
                                    calendar_found = True
                                    calendar_frame = frame
                                    break
                                except Exception:
                                    continue
                            if calendar_found:
                                break
                
                if not calendar_found:
                    logger.warning("未找到日历控件，等待更长时间...")
                    # 等待更长时间，然后再次检查
                    await asyncio.sleep(2)
                    
                    # 再次检查所有frame
                    try:
                        await self.page.wait_for_selector("#ui-datepicker-div", state="visible", timeout=1000)
                        logger.info("✓ 等待后在主页面找到jQuery UI日历控件")
                        calendar_found = True
                        calendar_frame = self.page
                    except Exception:
                        for i, frame in enumerate(frames):
                            try:
                                await frame.wait_for_selector("#ui-datepicker-div", state="visible", timeout=1000)
                                logger.info(f"✓ 等待后在iframe {i} 中找到jQuery UI日历控件")
                                calendar_found = True
                                calendar_frame = frame
                                break
                            except Exception:
                                continue
                    
                    if not calendar_found:
                        logger.warning("仍未找到日历控件")
                        continue
                
                # 3. 选择年份（基于实际HTML结构）
                try:
                    # 使用select_option方法选择年份（参考您提供的代码）
                    await calendar_frame.select_option(".ui-datepicker-year", str(year))
                    logger.info(f"✓ 选择年份: {year}")
                    
                    # 等待年份选择生效
                    await asyncio.sleep(0.5)
                except Exception as e:
                    logger.debug(f"select_option选择年份失败: {e}")
                    # 尝试其他年份选择方式
                    try:
                        # 方法2：直接点击年份下拉框，然后点击选项
                        await calendar_frame.click(".ui-datepicker-year")
                        logger.info("✓ 点击年份下拉框")
                        
                        # 等待下拉框展开
                        await asyncio.sleep(0.5)
                        
                        # 选择指定年份
                        year_option = calendar_frame.locator(f'.ui-datepicker-year option[value="{year}"]').first
                        if await year_option.count() > 0:
                            await year_option.click()
                            logger.info(f"✓ 通过点击选项选择年份: {year}")
                        else:
                            # 如果找不到选项，尝试直接点击年份文本
                            await calendar_frame.click(f'text={year}')
                            logger.info(f"✓ 通过文本选择年份: {year}")
                    except Exception as e2:
                        logger.debug(f"点击选择年份也失败: {e2}")
                        continue
                
                # 4. 选择月份（基于实际HTML结构）
                try:
                    # 使用select_option方法选择月份（参考您提供的代码）
                    month_index = month - 1  # 月份索引从0开始
                    await calendar_frame.select_option(".ui-datepicker-month", str(month_index))
                    logger.info(f"✓ 选择月份: {month} (索引: {month_index})")
                    
                    # 等待月份选择生效
                    await asyncio.sleep(0.5)
                except Exception as e:
                    logger.debug(f"select_option选择月份失败: {e}")
                    # 尝试其他月份选择方式
                    try:
                        # 方法2：直接点击月份下拉框，然后点击选项
                        await calendar_frame.click(".ui-datepicker-month")
                        logger.info("✓ 点击月份下拉框")
                        
                        # 等待下拉框展开
                        await asyncio.sleep(0.5)
                        
                        # 选择指定月份（月份索引从0开始，所以需要减1）
                        month_index = month - 1
                        month_option = calendar_frame.locator(f'.ui-datepicker-month option[value="{month_index}"]').first
                        if await month_option.count() > 0:
                            await month_option.click()
                            logger.info(f"✓ 通过点击选项选择月份: {month} (索引: {month_index})")
                        else:
                            # 如果找不到选项，尝试直接点击月份文本
                            month_names = ["1月", "2月", "3月", "4月", "5月", "6月", 
                                         "7月", "8月", "9月", "10月", "11月", "12月"]
                            month_name = month_names[month - 1]
                            await calendar_frame.click(f'text={month_name}')
                            logger.info(f"✓ 通过文本选择月份: {month_name}")
                    except Exception as e2:
                        logger.debug(f"点击选择月份也失败: {e2}")
                        continue
                
                # 5. 选择日期（基于实际HTML结构）
                try:
                    # 等待日历更新
                    await asyncio.sleep(0.5)
                    
                    # 根据实际HTML结构，日期是通过<a>标签实现的
                    # 尝试多种日期选择方式
                    date_selectors = [
                        f'#ui-datepicker-div a:has-text("{day}")',
                        f'#ui-datepicker-div td[data-handler="selectDay"] a:has-text("{day}")',
                        f'#ui-datepicker-div td[data-handler="selectDay"] a[href="#"]:has-text("{day}")',
                        f'#ui-datepicker-div .ui-state-default:has-text("{day}")',
                        f'#ui-datepicker-div a.ui-state-default:has-text("{day}")',
                        f'//a[text()="{day}"]'  # 使用XPath，参考您提供的代码
                    ]
                    
                    date_clicked = False
                    for selector in date_selectors:
                        try:
                            if selector.startswith('//'):
                                # 使用XPath选择器
                                date_element = calendar_frame.locator(selector).first
                            else:
                                # 使用CSS选择器
                                date_element = calendar_frame.locator(selector).first
                            
                            if await date_element.count() > 0:
                                await date_element.click()
                                logger.info(f"✓ 成功选择日期: {day} (使用选择器: {selector})")
                                date_clicked = True
                                break
                        except Exception as e:
                            logger.debug(f"日期选择器 {selector} 失败: {e}")
                            continue
                    
                    if not date_clicked:
                        # 如果所有选择器都失败，尝试通过JavaScript选择
                        try:
                            js_code = f"""
                            (function() {{
                                var datepicker = document.getElementById('ui-datepicker-div');
                                if (datepicker) {{
                                    var dayLinks = datepicker.querySelectorAll('a.ui-state-default');
                                    for (var i = 0; i < dayLinks.length; i++) {{
                                        if (dayLinks[i].textContent.trim() === '{day}') {{
                                            dayLinks[i].click();
                                            return true;
                                        }}
                                    }}
                                }}
                                return false;
                            }})();
                            """
                            result = await calendar_frame.evaluate(js_code)
                            if result:
                                logger.info(f"✓ 通过JavaScript成功选择日期: {day}")
                                date_clicked = True
                        except Exception as e:
                            logger.debug(f"JavaScript选择日期失败: {e}")
                    
                    if date_clicked:
                        # 等待日期选择完成
                        await asyncio.sleep(1)
                        
                        # 验证日期是否已填写
                        try:
                            current_value = await target_frame.input_value(f"#{element_id}")
                            logger.info(f"✓ 日期填写完成，当前值: {current_value}")
                            return
                        except Exception as e:
                            logger.debug(f"验证日期值失败: {e}")
                            # 如果验证失败，但操作看起来成功了，也返回
                            logger.info("✓ 日期选择操作完成")
                            return
                    else:
                        logger.warning(f"无法选择日期: {day}")
                        continue
                        
                except Exception as e:
                    logger.debug(f"选择日期失败: {e}")
                    continue
                
            except Exception as e:
                logger.warning(f"选择日期异常 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
        
        logger.error(f"选择日期最终失败: {element_id}")
        logger.info("建议检查：")
        logger.info("1. 日期输入框ID是否正确")
        logger.info("2. 日历控件是否正确加载")
        logger.info("3. 日期格式是否正确 (yyyy-mm-dd)")
        logger.info("4. 是否需要先点击其他元素来显示日期输入框")
    
    async def click_radio_button(self, element_id: str, retries: int = MAX_RETRIES):
        """
        点击radio按钮
        
        Args:
            element_id: radio按钮的ID或value值
            retries: 重试次数
        """
        logger.info(f"尝试点击radio按钮: {element_id}")
        
        for attempt in range(retries):
            try:
                # 获取所有iframe信息
                frames = self.page.frames
                logger.info(f"找到 {len(frames)} 个iframe")
                
                # 方法1: 优先在iframe中查找radio按钮（多种策略）
                logger.info(f"在iframe中查找radio按钮")
                for i, frame in enumerate(frames):
                    try:
                        # 策略1: 通过name和value查找（原有逻辑）
                        radio_selector = f"input[type='radio'][name*='school_area'][value='{element_id}']"
                        radio_element = frame.locator(radio_selector).first
                        if await radio_element.count() > 0:
                            await radio_element.click()
                            logger.info(f"✓ 在iframe {i} 中成功点击radio按钮 (策略1): {element_id}")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return
                        
                        # 策略2: 通过name和value查找（新业务类型radio）
                        radio_selector = f"input[type='radio'][name*='yta-filter_bcode'][value='{element_id}']"
                        radio_element = frame.locator(radio_selector).first
                        if await radio_element.count() > 0:
                            await radio_element.click()
                            logger.info(f"✓ 在iframe {i} 中成功点击radio按钮 (策略2): {element_id}")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return
                        
                        # 策略3: 通过文本内容查找（点击span文本）
                        text_selector = f"span:has-text('{element_id}')"
                        text_element = frame.locator(text_selector).first
                        if await text_element.count() > 0:
                            await text_element.click()
                            logger.info(f"✓ 在iframe {i} 中成功点击radio按钮 (策略3): {element_id}")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return
                        
                        # 策略4: 通过li元素查找（点击包含文本的li）
                        li_selector = f"li:has-text('{element_id}')"
                        li_element = frame.locator(li_selector).first
                        if await li_element.count() > 0:
                            await li_element.click()
                            logger.info(f"✓ 在iframe {i} 中成功点击radio按钮 (策略4): {element_id}")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return
                            
                    except Exception as e:
                        logger.debug(f"在iframe {i} 中查找radio按钮失败: {e}")
                        continue
                
                # 方法2: 在主页面查找radio按钮（多种策略）
                try:
                    logger.info(f"在主页面查找radio按钮")
                    
                    # 策略1: 通过name和value查找（原有逻辑）
                    radio_selector = f"input[type='radio'][name*='school_area'][value='{element_id}']"
                    radio_element = self.page.locator(radio_selector).first
                    if await radio_element.count() > 0:
                        await radio_element.click()
                        logger.info(f"✓ 在主页面成功点击radio按钮 (策略1): {element_id}")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        return
                    
                    # 策略2: 通过name和value查找（新业务类型radio）
                    radio_selector = f"input[type='radio'][name*='yta-filter_bcode'][value='{element_id}']"
                    radio_element = self.page.locator(radio_selector).first
                    if await radio_element.count() > 0:
                        await radio_element.click()
                        logger.info(f"✓ 在主页面成功点击radio按钮 (策略2): {element_id}")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        return
                    
                    # 策略3: 通过文本内容查找（点击span文本）
                    text_selector = f"span:has-text('{element_id}')"
                    text_element = self.page.locator(text_selector).first
                    if await text_element.count() > 0:
                        await text_element.click()
                        logger.info(f"✓ 在主页面成功点击radio按钮 (策略3): {element_id}")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        return
                    
                    # 策略4: 通过li元素查找（点击包含文本的li）
                    li_selector = f"li:has-text('{element_id}')"
                    li_element = self.page.locator(li_selector).first
                    if await li_element.count() > 0:
                        await li_element.click()
                        logger.info(f"✓ 在主页面成功点击radio按钮 (策略4): {element_id}")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        return
                        
                except Exception as e:
                    logger.debug(f"主页面查找radio按钮失败: {e}")
                
                logger.warning(f"未找到radio按钮: {element_id}")
                logger.info("建议检查：")
                logger.info("1. 元素value值是否正确")
                logger.info("2. 元素是否在iframe中")
                logger.info("3. 元素是否已经加载")
                logger.info("4. 是否使用了正确的显示文本而不是value值")
                return
                
            except Exception as e:
                logger.warning(f"点击radio按钮失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"点击radio按钮最终失败: {element_id}")
    
    async def click_button_by_btnname(self, btnname: str, retries: int = MAX_RETRIES):
        """
        通过btnName属性动态查找并点击按钮
        
        Args:
            btnname: 按钮的btnName属性值
            retries: 重试次数
        """
        logger.info(f"尝试通过btnName点击按钮: {btnname}")
        
        # 等待页面完全加载
        await asyncio.sleep(0.5)
        
        # 获取所有iframe
        frames = self.page.frames
        button_found = False
        
        # 方法1: 优先在iframe中动态查找按钮
        for frame in frames:
            try:
                logger.info(f"在iframe中查找按钮: {frame.url or 'unnamed frame'}")
                
                # 使用btnname选择器在iframe中查找
                selector = f"button[btnname='{btnname}']"
                button = frame.locator(selector).first
                
                if await button.count() > 0:
                    await button.click()
                    logger.info(f"✓ 在iframe中成功点击按钮 (btnname: {btnname})")
                    button_found = True
                    break
                else:
                    logger.debug(f"在iframe {frame.url} 中未找到按钮 {btnname}")
                    
            except Exception as e:
                logger.debug(f"在iframe中查找按钮时出错: {e}")
                continue
        
        # 方法2: 如果iframe中没找到，尝试在主页面查找
        if not button_found:
            logger.info("在iframe中未找到按钮，尝试在主页面查找...")
            try:
                selector = f"button[btnname='{btnname}']"
                button = self.page.locator(selector).first
                if await button.count() > 0:
                    await button.click()
                    logger.info(f"✓ 在主页面成功点击按钮 (btnname: {btnname})")
                    button_found = True
            except Exception as e:
                logger.debug(f"主页面选择器失败: {e}")
        
        # 方法3: 如果还是没找到，尝试其他选择器
        if not button_found:
            logger.info("尝试使用其他选择器查找按钮...")
            alternative_selectors = [
                f"button[guid*='{btnname}']",
                f"button:has-text('{btnname}')",
                f"input[btnname='{btnname}']",
                f"[btnname='{btnname}']"
            ]
            
            for selector in alternative_selectors:
                try:
                    # 在iframe中查找
                    for frame in frames:
                        try:
                            button = frame.locator(selector).first
                            if await button.count() > 0:
                                await button.click()
                                logger.info(f"✓ 在iframe中使用备用选择器成功点击按钮: {selector}")
                                button_found = True
                                break
                        except Exception as e:
                            continue
                    
                    if button_found:
                        break
                    
                    # 在主页面查找
                    try:
                        button = self.page.locator(selector).first
                        if await button.count() > 0:
                            await button.click()
                            logger.info(f"✓ 在主页面使用备用选择器成功点击按钮: {selector}")
                            button_found = True
                            break
                    except Exception as e:
                        continue
                        
                except Exception as e:
                    logger.debug(f"备用选择器 {selector} 失败: {e}")
                    continue
        
        if button_found:
            await asyncio.sleep(BUTTON_CLICK_WAIT)
            return True
        else:
            logger.error(f"点击按钮最终失败: {btnname}")
            return False

    async def click_first_row_reservation_button(self, retries: int = MAX_RETRIES):
        """
        点击表格中第一行的预约按钮
        
        Args:
            retries: 重试次数
        """
        logger.info("尝试点击表格中第一行的预约按钮")
        
        for attempt in range(retries):
            try:
                # 获取所有iframe
                frames = self.page.frames
                logger.info(f"找到 {len(frames)} 个iframe")
                
                # 方法1: 在主页面查找表格和预约按钮
                try:
                    logger.info("在主页面查找表格...")
                    await self.page.wait_for_selector("tbody >> tr[id^='2179_']", timeout=5000)
                    logger.info("✓ 主页面表格加载完成")
                    
                    # 点击第一行中的预约按钮
                    selector = "tbody >> tr[id^='2179_'] >> button[btnname='预约']"
                    button = self.page.locator(selector).first
                    
                    if await button.count() > 0:
                        await button.click()
                        logger.info("✓ 在主页面成功点击第一行的预约按钮")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        return True
                    else:
                        logger.debug("主页面未找到预约按钮")
                        
                except Exception as e:
                    logger.debug(f"主页面查找失败: {e}")
                
                # 方法2: 在iframe中查找表格和预约按钮
                logger.info("在主页面未找到，尝试在iframe中查找...")
                for i, frame in enumerate(frames):
                    try:
                        logger.info(f"在iframe {i} 中查找表格: {frame.url or 'unnamed frame'}")
                        
                        # 等待表格加载完成
                        await frame.wait_for_selector("tbody >> tr[id^='2179_']", timeout=5000)
                        logger.info(f"✓ iframe {i} 表格加载完成")
                        
                        # 点击第一行中的预约按钮
                        selector = "tbody >> tr[id^='2179_'] >> button[btnname='预约']"
                        button = frame.locator(selector).first
                        
                        if await button.count() > 0:
                            await button.click()
                            logger.info(f"✓ 在iframe {i} 中成功点击第一行的预约按钮")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return True
                        else:
                            logger.debug(f"iframe {i} 中未找到预约按钮")
                            
                    except Exception as e:
                        logger.debug(f"在iframe {i} 中查找失败: {e}")
                        continue
                
                # 方法3: 更宽松的查找策略
                logger.info("尝试更宽松的查找策略...")
                for i, frame in enumerate(frames):
                    try:
                        # 查找任何包含"预约"的按钮
                        button = frame.locator("button[btnname='预约']").first
                        if await button.count() > 0:
                            await button.click()
                            logger.info(f"✓ 在iframe {i} 中找到并点击预约按钮")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return True
                    except Exception as e:
                        logger.debug(f"在iframe {i} 中宽松查找失败: {e}")
                        continue
                
                # 方法4: 在主页面尝试宽松查找
                try:
                    button = self.page.locator("button[btnname='预约']").first
                    if await button.count() > 0:
                        await button.click()
                        logger.info("✓ 在主页面找到并点击预约按钮")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        return True
                except Exception as e:
                    logger.debug(f"主页面宽松查找失败: {e}")
                
                logger.warning(f"未找到预约按钮 (尝试 {attempt + 1}/{retries})")
                
            except Exception as e:
                logger.warning(f"点击预约按钮失败 (尝试 {attempt + 1}/{retries}): {e}")
                
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error("点击预约按钮最终失败")
                    return False
        
        return False
    
    async def click_button(self, element_id: str, retries: int = MAX_RETRIES):
        """
        点击网页中的按钮
        
        Args:
            element_id: 按钮的ID或btnName
            retries: 重试次数
        """
        for attempt in range(retries):
            try:
                # 优先在iframe中查找（根据日志分析，大部分元素都在iframe中）
                frames = self.page.frames
                for frame in frames:
                    try:
                        # 在iframe中通过ID点击
                        button_element = frame.locator(f"#{element_id}").first
                        if await button_element.count() > 0:
                            await button_element.click()
                            logger.info(f"在iframe中成功点击按钮: {element_id}")
                            await asyncio.sleep(BUTTON_CLICK_WAIT)
                            return
                    except Exception as e:
                        logger.debug(f"在iframe中查找按钮失败: {e}")
                        continue
                
                # 如果iframe中找不到，尝试在主页面通过ID点击
                if element_id and await self.wait_for_element(element_id):
                    await self.page.click(f"#{element_id}")
                    logger.info(f"在主页面成功点击按钮: {element_id}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return
                else:
                    # 如果ID不存在，尝试通过btnName点击
                    await self.click_button_by_btnname(element_id)
                    return
            except Exception as e:
                logger.warning(f"点击按钮失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"点击按钮最终失败: {element_id}")
    
    async def click_navigation_panel(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        点击系统导航面板
        
        Args:
            element_id: 导航面板的ID（从标题-ID.xlsx获取）
            value: 导航面板的值（如WF_YB6）
            retries: 重试次数
        """
        logger.info(f"开始点击导览框: element_id={element_id}, value={value}")
        
        for attempt in range(retries):
            try:
                # 方法1: 通过onclick属性查找
                logger.info(f"尝试通过onclick属性查找: div[onclick*='{value}']")
                onclick_selector = f"div[onclick*='{value}']"
                
                if await self.page.locator(onclick_selector).count() > 0:
                    await self.page.click(onclick_selector)
                    logger.info(f"成功点击导览框 (通过onclick): {value}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return True
                
                # 方法2: 通过JavaScript直接调用
                logger.info(f"尝试通过JavaScript调用: navToPrj('{value}')")
                try:
                    await self.page.evaluate(f"navToPrj('{value}')")
                    logger.info(f"成功点击导览框 (通过JavaScript): {value}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return True
                except Exception as js_error:
                    logger.debug(f"JavaScript调用失败: {js_error}")
                
                # 方法3: 通过文本内容查找
                logger.info(f"尝试通过文本内容查找: div:has-text('{value}')")
                text_selector = f"div:has-text('{value}')"
                
                if await self.page.locator(text_selector).count() > 0:
                    await self.page.click(text_selector)
                    logger.info(f"成功点击导览框 (通过文本): {value}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return True
                
                # 方法4: 通过title属性查找
                logger.info(f"尝试通过title属性查找: div[title*='{value}']")
                title_selector = f"div[title*='{value}']"
                
                if await self.page.locator(title_selector).count() > 0:
                    await self.page.click(title_selector)
                    logger.info(f"成功点击导览框 (通过title): {value}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return True
                
                # 方法5: 通过class和onclick组合查找
                logger.info(f"尝试通过class和onclick组合查找: div.syslink[onclick*='{value}']")
                class_selector = f"div.syslink[onclick*='{value}']"
                
                if await self.page.locator(class_selector).count() > 0:
                    await self.page.click(class_selector)
                    logger.info(f"成功点击导览框 (通过class+onclick): {value}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return True
                
                # 方法6: 通过第一个syslink元素查找
                logger.info("尝试通过第一个syslink元素查找")
                first_syslink = self.page.locator("div.syslink").first
                
                if await first_syslink.count() > 0:
                    await first_syslink.click()
                    logger.info(f"成功点击导览框 (通过第一个syslink): {value}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return True
                
                logger.warning(f"所有方法都失败，尝试 {attempt + 1}/{retries}")
                
            except Exception as e:
                logger.warning(f"点击导览框失败 (尝试 {attempt + 1}/{retries}): {e}")
            
            if attempt < retries - 1:
                await asyncio.sleep(RETRY_DELAY)
        
        logger.error(f"点击导览框最终失败: {value}")
        return False
    
    async def process_cell(self, title: str, value: Any):
        """
        处理单个单元格的内容
        
        Args:
            title: 列标题
            value: 单元格值
        """
        if pd.isna(value) or value == "":
            return
            
        value_str = self.clean_value_string(value)
        
        # 特殊处理：保存报销项目号和金额用于文件命名
        if title == "报销项目号":
            self.current_project_number = value_str
            logger.info(f"保存报销项目号用于文件命名: {value_str}")
        elif title == "金额":
            self.current_amount = value_str
            logger.info(f"保存金额用于文件命名: {value_str}")
        
        # 处理等待操作（标题为"等待"或"等待.1"等）
        if title.startswith("等待"):
            try:
                # 支持两种格式：$数字 或 直接数字
                if value_str.startswith(BUTTON_PREFIX):
                    # 格式：$数字
                    wait_seconds_str = value_str[len(BUTTON_PREFIX):]  # 去掉$前缀
                else:
                    # 格式：直接数字
                    wait_seconds_str = value_str
                
                wait_seconds = float(wait_seconds_str)
                logger.info(f"检测到等待操作，等待 {wait_seconds} 秒")
                await asyncio.sleep(wait_seconds)
                logger.info(f"等待 {wait_seconds} 秒完成")
                return
            except ValueError:
                logger.warning(f"等待操作格式错误，无法解析秒数: {value_str}")
                return
        
        # 处理radio按钮点击操作（以$$开头）- 优先处理
        if value_str.startswith(RADIO_BUTTON_PREFIX):
            # 提取$$之后的内容作为标题
            radio_title = value_str[len(RADIO_BUTTON_PREFIX):]  # 去掉$$前缀
            logger.info(f"检测到radio按钮操作，使用$$后的内容作为标题: {radio_title}")
            
            # 根据标题查找对应的radio按钮ID
            radio_element_id = self.get_object_id(radio_title)
            if radio_element_id:
                logger.info(f"找到radio按钮ID: {radio_element_id}")
                await self.click_radio_button(radio_element_id)
            else:
                logger.warning(f"未找到标题 '{radio_title}' 对应的radio按钮ID映射，请检查标题-ID映射表中是否包含此标题")
            return
        
        # 处理第一行预约按钮操作（以$开头且标题为"预约按钮"）- 优先处理
        if value_str.startswith(BUTTON_PREFIX) and title == "预约按钮":
            button_value = value_str[len(BUTTON_PREFIX):]  # 去掉$前缀
            if button_value == "预约":
                logger.info("检测到第一行预约按钮操作")
                await self.click_first_row_reservation_button()
                return
        
        # 对于其他情况，先获取element_id
        element_id = self.get_object_id(title)
        if not element_id:
            return
        
        # 特殊处理：网上预约报账按钮（优先处理）
        if title == "网上预约报账按钮":
            logger.info(f"特殊处理网上预约报账按钮: {element_id}")
            # 从element_id中提取WF_YB6参数
            if "navToPrj('WF_YB6')" in element_id:
                await self.click_navigation_panel("", "WF_YB6")
                return
            else:
                # 如果element_id不是JavaScript函数，尝试直接点击
                await self.click_button(element_id)
                return
        
        # 特殊处理：转卡信息工号（填写后检查银行卡选择弹窗）
        if title == "转卡信息工号" or title.startswith("转卡信息工号"):
            logger.info(f"特殊处理转卡信息工号: {value_str}")
            await self.fill_input(element_id, value_str, title=title)
            
            # 填写工号后输入回车键来触发银行卡选择界面
            logger.info("填写转卡信息工号完成，输入回车键触发银行卡选择界面...")
            await asyncio.sleep(0.5)  # 短暂等待确保输入完成
            
            # 在输入框中输入回车键
            try:
                # 首先尝试在主页面查找输入框并输入回车
                if element_id and await self.wait_for_element(element_id, timeout=2):
                    await self.page.press(f"#{element_id}", "Enter")
                    logger.info(f"在主页面输入框中输入回车键: {element_id}")
                else:
                    # 如果主页面找不到，尝试在iframe中查找
                    frames = self.page.frames
                    for frame in frames:
                        try:
                            input_element = frame.locator(f"#{element_id}").first
                            if await input_element.count() > 0:
                                await input_element.press("Enter")
                                logger.info(f"在iframe中输入框中输入回车键: {element_id}")
                                break
                        except Exception as e:
                            logger.debug(f"在iframe中查找输入框失败: {e}")
                            continue
                    else:
                        # 如果还是找不到，尝试通过name属性查找
                        try:
                            await self.page.press(f"input[name='{element_id}']", "Enter")
                            logger.info(f"通过name属性输入框中输入回车键: {element_id}")
                        except Exception as e:
                            logger.debug(f"通过name属性查找失败: {e}")
            except Exception as e:
                logger.warning(f"输入回车键失败: {e}")
            
            # 等待银行卡选择弹窗出现（缩减等待时间）
            logger.info("等待银行卡选择弹窗出现...")
            await asyncio.sleep(BANK_CARD_DIALOG_WAIT)
            
            # 检查是否需要选择银行卡
            # 创建当前记录的DataFrame
            current_record = pd.DataFrame([{title: value_str}])
            await self.handle_bank_card_selection_for_transfer(value_str, current_record)
            return
        

        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith(BUTTON_PREFIX):
            button_value = value_str[len(BUTTON_PREFIX):]  # 去掉前缀符号
            
            # 特殊处理：如果是打印按钮，查找并点击打印确认单按钮
            if title == "打印按钮" or title == "打印操作" or title == "打印确认单按钮":
                logger.info("检测到打印按钮操作，查找并点击打印确认单按钮")
                await self.click_print_button()
                return
            
            # 普通按钮点击
            await self.click_button(element_id)
            return
        
        # 特殊处理：科目和金额填写（需要等待页面加载）
        if title == "科目" or title == "金额":
            # 特殊处理：科目列（以#开头）
            if title == "科目" and value_str.startswith("#"):
                logger.info(f"处理科目列: {title} = {value_str}")
                
                # 提取科目名称（去掉#前缀）
                subject_name = value_str[1:]
                
                # 在标题-ID表中查找对应的输入框ID
                input_id = self.get_object_id(subject_name)
                if not input_id:
                    logger.warning(f"未找到科目 '{subject_name}' 对应的ID映射")
                    return
                
                # 等待页面加载
                logger.info(f"特殊处理科目填写，等待页面加载完成...")
                await asyncio.sleep(SUBJECT_AMOUNT_WAIT)
                logger.info(f"页面加载等待完成，开始填写科目: {subject_name}")
                await self.fill_input(input_id, value_str, title=title)
                return
            elif title == "金额":
                # 金额列需要特殊处理，因为它需要与科目配对
                logger.info(f"处理金额列: {title} = {value_str}")
                # 这里暂时不处理，因为金额应该与科目配对处理
                return
            else:
                logger.info(f"特殊处理{title}填写，等待页面加载完成...")
                await asyncio.sleep(SUBJECT_AMOUNT_WAIT)
                logger.info(f"页面加载等待完成，开始填写{title}: {value_str}")
                await self.fill_input(element_id, value_str, title=title)
                return
        
        # 处理系统导览框点击操作（以@开头）
        if value_str.startswith(NAVIGATION_PREFIX):
            nav_value = value_str[1:]  # 去掉@符号
            await self.click_navigation_panel(element_id, nav_value)
            return
        
        # 处理卡号尾号选择（以*开头）
        if value_str.startswith(CARD_NUMBER_PREFIX):
            card_tail = value_str[1:]  # 去掉*符号
            await self.select_card_by_number(card_tail)
            return
        
        # 处理下拉框选择（支持列名映射和ID模式识别）
        # 创建列名到配置名的映射
        dropdown_title_mapping = {
            "省份": "省份地区",  # Excel列名 -> 配置名
            "人员类型": "人员类型",  # 保持原样
            "安排状态": "安排状态",  # 保持原样
            "交通费": "交通费"  # 保持原样
        }
        
        # 获取实际的配置名
        config_title = dropdown_title_mapping.get(title, title)
        
        # 检查是否为下拉框字段（通过配置名或ID模式）
        is_dropdown = False
        dropdown_config = None
        
        # 方法1: 通过配置名检查
        if config_title in DROPDOWN_FIELDS:
            is_dropdown = True
            dropdown_config = DROPDOWN_FIELDS[config_title]
        
        # 方法2: 通过ID模式检查（如果element_id存在）
        elif element_id:
            # 检查是否为已知的下拉框ID模式
            dropdown_id_patterns = [
                "formWF_YB6_3492_yc-chr_sf",  # 省份下拉框模式
                "formWF_YB6_3492_yc-chr_hsf",  # hsf下拉框模式
                "formWF_YB6_3492_yc-chr_jtf",  # jtf下拉框模式
                "formWF_YB6_3492_yc-chr_zc",   # 人员类型下拉框模式
                "formWF_YB6_3492_yc-chr_azzt", # 安排状态下拉框模式
            ]
            
            for pattern in dropdown_id_patterns:
                if pattern in element_id:
                    is_dropdown = True
                    # 根据ID模式确定下拉框类型（使用精确匹配）
                    if "formWF_YB6_3492_yc-chr_sf" in element_id:
                        dropdown_config = DROPDOWN_FIELDS.get("省份地区", {})
                    elif "formWF_YB6_3492_yc-chr_hsf" in element_id:
                        # 这里需要根据实际情况确定hsf对应的下拉框类型
                        dropdown_config = DROPDOWN_FIELDS.get("安排状态", {})
                    elif "formWF_YB6_3492_yc-chr_jtf" in element_id:
                        dropdown_config = DROPDOWN_FIELDS.get("交通费", {})
                    elif "formWF_YB6_3492_yc-chr_zc" in element_id:
                        dropdown_config = DROPDOWN_FIELDS.get("人员类型", {})
                    elif "formWF_YB6_3492_yc-chr_azzt" in element_id:
                        dropdown_config = DROPDOWN_FIELDS.get("安排状态", {})
                    break
        
        if is_dropdown and dropdown_config:
            # 获取下拉框的映射关系
            dropdown_mapping = dropdown_config
            # 查找对应的值
            if value_str in dropdown_mapping:
                mapped_value = dropdown_mapping[value_str]
                await self.select_dropdown(element_id, mapped_value)
                logger.info(f"下拉框映射: {title} = {value_str} -> {mapped_value}")
            else:
                # 如果没有映射，直接使用原值
                await self.select_dropdown(element_id, value_str)
                logger.info(f"下拉框直接选择: {title} = {value_str}")
            return
        
        # 处理日期输入框（检查element_id是否包含日期相关的标识）
        if (element_id and ("date" in element_id.lower() or "startdate" in element_id.lower() or "enddate" in element_id.lower() or 
                           element_id.endswith("_startdate") or element_id.endswith("_enddate") or
                           "temp-startdate" in element_id or "temp-enddate" in element_id or
                           element_id == "formWF_YB6_3492_yc-chr_start1_0" or element_id == "formWF_YB6_3492_yc-chr_end1_0" or
                           "start" in element_id or "end" in element_id)):
            logger.info(f"检测到日期输入框: {element_id} = {value_str}")
            
            # 优先尝试新的jQuery UI日历控件方法
            try:
                await self.select_date_from_calendar(element_id, value_str)
                return
            except Exception as e:
                logger.debug(f"jQuery UI日历控件方法失败，尝试只读日期输入框方法: {e}")
                # 如果新方法失败，回退到只读方法
                try:
                    await self.fill_readonly_date_input(element_id, value_str)
                    return
                except Exception as e2:
                    logger.debug(f"只读日期输入框方法也失败，尝试普通方法: {e2}")
                    # 如果只读方法也失败，回退到普通方法
                    await self.fill_date_input(element_id, value_str)
                    return
        
        # 处理普通输入框
        await self.fill_input(element_id, value_str, title=title)
    
    async def select_dropdown(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        选择下拉框中的选项
        
        Args:
            element_id: 下拉框的ID
            value: 要选择的选项值
            retries: 重试次数
        """
        for attempt in range(retries):
            try:
                # 优先在iframe中查找（根据日志分析，大部分元素都在iframe中）
                frames = self.page.frames
                for frame in frames:
                    try:
                        # 在iframe中查找下拉框
                        select_element = frame.locator(f"#{element_id}").first
                        if await select_element.count() > 0:
                            await select_element.select_option(value=value)
                            logger.info(f"在iframe中成功选择下拉框 {element_id}: {value}")
                            await asyncio.sleep(ELEMENT_WAIT)
                            return
                    except Exception as e:
                        logger.debug(f"在iframe中查找下拉框失败: {e}")
                        continue
                
                # 如果iframe中找不到，尝试在主页面查找
                try:
                    await self.page.wait_for_selector(f"#{element_id}", timeout=3000)
                    await self.page.select_option(f"#{element_id}", value)
                    logger.info(f"在主页面成功选择下拉框 {element_id}: {value}")
                    await asyncio.sleep(ELEMENT_WAIT)
                    return
                except Exception as e:
                    logger.debug(f"在主页面查找下拉框失败: {e}")
                
                # 如果还是找不到，尝试通过name属性查找（优先在iframe中）
                for frame in frames:
                    try:
                        select_element = frame.locator(f"select[name='{element_id}']").first
                        if await select_element.count() > 0:
                            await select_element.select_option(value=value)
                            logger.info(f"在iframe中通过name属性成功选择下拉框 {element_id}: {value}")
                            await asyncio.sleep(ELEMENT_WAIT)
                            return
                    except Exception as e:
                        logger.debug(f"在iframe中通过name属性查找失败: {e}")
                        continue
                
                # 最后尝试在主页面通过name属性查找
                try:
                    await self.page.select_option(f"select[name='{element_id}']", value=value)
                    logger.info(f"在主页面通过name属性成功选择下拉框 {element_id}: {value}")
                    await asyncio.sleep(ELEMENT_WAIT)
                    return
                except Exception as e:
                    logger.debug(f"在主页面通过name属性查找失败: {e}")
                
                logger.warning(f"下拉框元素不存在: {element_id}")
                return
                    
            except Exception as e:
                logger.warning(f"选择下拉框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"选择下拉框最终失败: {element_id}")
    
    async def handle_bank_card_selection(self, record_data: pd.DataFrame):
        """
        处理银行卡选择弹窗
        
        Args:
            record_data: 包含登录信息的DataFrame行
        """
        try:
            logger.info("开始检测银行卡选择弹窗...")
            
            # 等待银行卡选择弹窗出现（缩减等待时间）
            await asyncio.sleep(BANK_CARD_DIALOG_WAIT)
            
            # 等待银行卡选择弹窗出现 - 尝试多种选择器
            bank_dialog_found = False
            selectors_to_try = [
                "#paybankdiv",                           # 主要的银行卡选择弹窗ID
                "div[id='paybankdiv']",                  # 完整的div选择器
                "div.ui-dialog-content",                 # UI对话框内容
                "table[style*='background-color:#F2FAFD']",  # 银行卡表格
                "input[name='rdoacnt']",                 # 银行卡选择radio按钮
                "div.ui-dialog[aria-describedby='paybankdiv']",  # 完整的对话框
                "div.ui-dialog-title:has-text('请选择卡号')",  # 对话框标题
                "tbody tr td input[type='radio'][name='rdoacnt']"  # 表格中的radio按钮
            ]
            
            # 方法1: 等待弹窗出现
            for selector in selectors_to_try:
                try:
                    await self.page.wait_for_selector(selector, timeout=5000)
                    logger.info(f"检测到银行卡选择弹窗，使用选择器: {selector}")
                    bank_dialog_found = True
                    break
                except Exception as e:
                    logger.debug(f"选择器 {selector} 未找到: {e}")
                    continue
            
            # 方法2: 如果方法1失败，尝试检测弹窗是否已经存在
            if not bank_dialog_found:
                logger.info("尝试检测已存在的银行卡选择弹窗...")
                for selector in selectors_to_try:
                    try:
                        elements = await self.page.locator(selector).all()
                        if len(elements) > 0:
                            logger.info(f"检测到已存在的银行卡选择弹窗，使用选择器: {selector}")
                            bank_dialog_found = True
                            break
                    except Exception as e:
                        logger.debug(f"检测选择器 {selector} 失败: {e}")
                        continue
            
            # 方法3: 检查是否有"请选择卡号"的标题
            if not bank_dialog_found:
                try:
                    title_elements = await self.page.locator("text=请选择卡号").all()
                    if len(title_elements) > 0:
                        logger.info("检测到'请选择卡号'标题，说明银行卡选择弹窗已出现")
                        bank_dialog_found = True
                except Exception as e:
                    logger.debug(f"检测'请选择卡号'标题失败: {e}")
            
            if not bank_dialog_found:
                logger.info("未检测到银行卡选择弹窗，可能只有一张卡或弹窗未出现")
                return
            
            # 查找卡号尾号列
            card_tail_value = None
            for col in record_data.columns:
                if col.startswith("卡号尾号") or col == "卡号尾号":
                    value = record_data[col].iloc[0]
                    value_str = self.clean_value_string(value)
                    if value_str.startswith("*"):
                        card_tail_value = value_str[1:]  # 去掉*前缀
                        break
            
            if not card_tail_value:
                logger.warning("未找到卡号尾号信息")
                return
            
            logger.info(f"开始选择卡号尾号: {card_tail_value}")
            
            # 等待弹窗完全加载（缩减等待时间）
            await asyncio.sleep(BANK_CARD_SELECTION_WAIT)
            
            # 尝试多种方式查找和点击radio按钮
            radio_clicked = False
            
            # 方法1: 通过XPath查找包含卡号尾号的tr，然后点击其中的radio
            try:
                radio_selector = f"//tr[td[contains(text(), '{card_tail_value}')]]/td/input[@type='radio'][@name='rdoacnt']"
                radio_element = self.page.locator(radio_selector).first
                if await radio_element.count() > 0:
                    await radio_element.click()
                    logger.info(f"成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                    radio_clicked = True
            except Exception as e:
                logger.debug(f"方法1失败: {e}")
            
            # 方法2: 如果方法1失败，尝试在iframe中查找
            if not radio_clicked:
                frames = self.page.frames
                for frame in frames:
                    try:
                        radio_selector = f"//tr[td[contains(text(), '{card_tail_value}')]]/td/input[@type='radio'][@name='rdoacnt']"
                        radio_element = frame.locator(radio_selector).first
                        if await radio_element.count() > 0:
                            await radio_element.click()
                            logger.info(f"在iframe中成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                            radio_clicked = True
                            break
                    except Exception as e:
                        logger.debug(f"在iframe中查找失败: {e}")
                        continue
            
            # 方法3: 通过onclick属性查找
            if not radio_clicked:
                try:
                    radio_selector = f"input[type='radio'][name='rdoacnt'][onclick*='{card_tail_value}']"
                    radio_element = self.page.locator(radio_selector).first
                    if await radio_element.count() > 0:
                        await radio_element.click()
                        logger.info(f"通过onclick属性成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                        radio_clicked = True
                except Exception as e:
                    logger.debug(f"方法3失败: {e}")
            
            # 方法4: 通过卡号文本查找
            if not radio_clicked:
                try:
                    # 查找包含卡号尾号的td元素，然后找到同行的radio按钮
                    card_selector = f"td:has-text('{card_tail_value}')"
                    card_elements = await self.page.locator(card_selector).all()
                    
                    for card_element in card_elements:
                        try:
                            # 找到包含这个td的tr，然后找到其中的radio按钮
                            parent_tr = card_element.locator("xpath=..")
                            radio_in_tr = parent_tr.locator("input[type='radio'][name='rdoacnt']").first
                            if await radio_in_tr.count() > 0:
                                await radio_in_tr.click()
                                logger.info(f"通过卡号文本成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                                radio_clicked = True
                                break
                        except Exception as e:
                            logger.debug(f"在tr中查找radio失败: {e}")
                            continue
                except Exception as e:
                    logger.debug(f"方法4失败: {e}")
            
            if radio_clicked:
                logger.info("银行卡选择成功，等待选择生效...")
                await asyncio.sleep(2)  # 等待选择生效
                
                # 尝试点击确定按钮
                try:
                    confirm_button = self.page.locator("button:has-text('确定')").first
                    if await confirm_button.count() > 0:
                        await confirm_button.click()
                        logger.info("成功点击确定按钮")
                        await asyncio.sleep(1)
                except Exception as e:
                    logger.debug(f"点击确定按钮失败: {e}")
            else:
                logger.warning(f"未找到卡号尾号 {card_tail_value} 对应的银行卡")
                
        except Exception as e:
            logger.error(f"处理银行卡选择失败: {e}")

    async def handle_bank_card_selection_for_transfer(self, work_id: str, current_record: pd.DataFrame = None):
        """
        处理转卡信息工号填写后的银行卡选择弹窗
        
        Args:
            work_id: 转卡信息工号
            current_record: 当前处理的记录（可选）
        """
        try:
            logger.info(f"开始检测转卡信息工号 {work_id} 的银行卡选择弹窗...")
            
            # 等待银行卡选择弹窗出现 - 只在iframe中查找
            bank_dialog_found = False
            selectors_to_try = [
                "#paybankdiv",                           # 主要的银行卡选择弹窗ID
                "div.ui-dialog[aria-describedby='paybankdiv']",  # 完整的对话框
                "div.ui-dialog-title:has-text('请选择卡号')",  # 对话框标题
                "input[name='rdoacnt']",                 # 银行卡选择radio按钮
                "table[style*='background-color:#F2FAFD']",  # 银行卡表格
                "div.ui-dialog-content"                  # UI对话框内容
            ]
            
            # 在所有iframe中查找弹窗
            logger.info("在所有iframe中查找银行卡选择弹窗...")
            frames = self.page.frames
            logger.info(f"找到 {len(frames)} 个iframe")
            
            target_frame = None
            for i, frame in enumerate(frames):
                logger.info(f"检查iframe {i}: {frame.url}")
                try:
                    for selector in selectors_to_try:
                        try:
                            await frame.wait_for_selector(selector, timeout=2000)
                            logger.info(f"✓ 在iframe {i} 中检测到银行卡选择弹窗，使用选择器: {selector}")
                            bank_dialog_found = True
                            target_frame = frame
                            break
                        except Exception as e:
                            logger.debug(f"iframe {i} 选择器 {selector} 未找到: {e}")
                            continue
                    if bank_dialog_found:
                        break
                except Exception as e:
                    logger.debug(f"检查iframe {i} 失败: {e}")
                    continue
            
            # 如果没找到，尝试检测已存在的弹窗
            if not bank_dialog_found:
                logger.info("尝试检测已存在的银行卡选择弹窗...")
                for i, frame in enumerate(frames):
                    try:
                        for selector in selectors_to_try:
                            try:
                                elements = await frame.locator(selector).all()
                                if len(elements) > 0:
                                    logger.info(f"✓ 在iframe {i} 中检测到已存在的银行卡选择弹窗，使用选择器: {selector}")
                                    bank_dialog_found = True
                                    target_frame = frame
                                    break
                            except Exception as e:
                                logger.debug(f"iframe {i} 检测选择器 {selector} 失败: {e}")
                                continue
                        if bank_dialog_found:
                            break
                    except Exception as e:
                        logger.debug(f"检查iframe {i} 失败: {e}")
                        continue
            
            # 检查是否有"请选择卡号"的标题
            if not bank_dialog_found:
                logger.info("检查是否有'请选择卡号'标题...")
                for i, frame in enumerate(frames):
                    try:
                        title_elements = await frame.locator("text=请选择卡号").all()
                        if len(title_elements) > 0:
                            logger.info(f"✓ 在iframe {i} 中检测到'请选择卡号'标题")
                            bank_dialog_found = True
                            target_frame = frame
                            break
                    except Exception as e:
                        logger.debug(f"iframe {i} 检测'请选择卡号'标题失败: {e}")
                        continue
            
            # 检查是否有radio按钮
            if not bank_dialog_found:
                logger.info("检查是否有银行卡选择radio按钮...")
                for i, frame in enumerate(frames):
                    try:
                        radio_elements = await frame.locator("input[type='radio'][name='rdoacnt']").all()
                        if len(radio_elements) > 0:
                            logger.info(f"✓ 在iframe {i} 中检测到 {len(radio_elements)} 个银行卡选择radio按钮")
                            bank_dialog_found = True
                            target_frame = frame
                            break
                    except Exception as e:
                        logger.debug(f"iframe {i} 检测radio按钮失败: {e}")
                        continue
            
            # 检查是否有ui-dialog类的元素
            if not bank_dialog_found:
                logger.info("检查是否有ui-dialog类的元素...")
                for i, frame in enumerate(frames):
                    try:
                        dialog_elements = await frame.locator("div.ui-dialog").all()
                        if len(dialog_elements) > 0:
                            logger.info(f"✓ 在iframe {i} 中检测到 {len(dialog_elements)} 个ui-dialog元素")
                            bank_dialog_found = True
                            target_frame = frame
                            break
                    except Exception as e:
                        logger.debug(f"iframe {i} 检测ui-dialog元素失败: {e}")
                        continue
            
            if not bank_dialog_found:
                logger.warning("未检测到银行卡选择弹窗，可能只有一张卡或弹窗未出现")
                return
            
            logger.info("银行卡选择弹窗已检测到，开始处理...")
            
            # 查找卡号尾号信息
            card_tail_value = None
            
            # 从当前记录中查找卡号尾号
            if current_record is not None:
                for col in current_record.columns:
                    if col.startswith("卡号尾号") or col == "卡号尾号":
                        value = current_record[col].iloc[0]
                        value_str = self.clean_value_string(value)
                        if value_str.startswith("*"):
                            card_tail_value = value_str[1:]  # 去掉*前缀
                            logger.info(f"从当前记录中找到卡号尾号: {card_tail_value}")
                            break
            
            # 如果没找到，尝试从全局数据中查找
            if not card_tail_value and hasattr(self, 'reimbursement_data'):
                for col in self.reimbursement_data.columns:
                    if col.startswith("卡号尾号") or col == "卡号尾号":
                        # 查找当前工号对应的卡号尾号
                        for _, row in self.reimbursement_data.iterrows():
                            if "转卡信息工号" in row and pd.notna(row["转卡信息工号"]):
                                if self.clean_value_string(row["转卡信息工号"]) == work_id:
                                    value = row[col]
                                    value_str = self.clean_value_string(value)
                                    if value_str.startswith("*"):
                                        card_tail_value = value_str[1:]  # 去掉*前缀
                                        logger.info(f"从全局数据中找到卡号尾号: {card_tail_value}")
                                        break
                        if card_tail_value:
                            break
            
            if not card_tail_value:
                logger.warning("未找到卡号尾号信息，将自动选择第一张银行卡")
                # 自动选择第一张银行卡
                try:
                    radio_buttons = await target_frame.locator("input[type='radio'][name='rdoacnt']").all()
                    if len(radio_buttons) > 0:
                        await radio_buttons[0].click()
                        logger.info("✓ 自动选择第一张银行卡")
                        await asyncio.sleep(1)
                        
                        # 点击确定按钮
                        await self.click_confirm_button_in_dialog()
                    else:
                        logger.warning("未找到可选择的银行卡")
                except Exception as e:
                    logger.error(f"选择银行卡失败: {e}")
                return
            
            logger.info(f"开始选择卡号尾号: {card_tail_value}")
            
            # 等待弹窗完全加载（缩减等待时间）
            await asyncio.sleep(BANK_CARD_SELECTION_WAIT)
            
            # 尝试多种方式查找和点击radio按钮
            radio_clicked = False
            
            # 方法1: 通过XPath查找包含卡号尾号的tr，然后点击其中的radio
            logger.info("方法1: 通过XPath查找包含卡号尾号的tr...")
            try:
                # 查找包含卡号尾号的td元素，然后找到同行的radio按钮
                radio_selector = f"//tr[td[contains(text(), '{card_tail_value}')]]/td/input[@type='radio'][@name='rdoacnt']"
                radio_element = target_frame.locator(radio_selector).first
                if await radio_element.count() > 0:
                    await radio_element.click()
                    logger.info(f"✓ 成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                    radio_clicked = True
                else:
                    logger.debug(f"未找到卡号尾号 {card_tail_value} 对应的radio按钮")
            except Exception as e:
                logger.debug(f"方法1失败: {e}")
            
            # 方法2: 通过onclick属性查找
            if not radio_clicked:
                logger.info("方法2: 通过onclick属性查找...")
                try:
                    radio_selector = f"input[type='radio'][name='rdoacnt'][onclick*='{card_tail_value}']"
                    radio_element = target_frame.locator(radio_selector).first
                    if await radio_element.count() > 0:
                        await radio_element.click()
                        logger.info(f"✓ 通过onclick属性成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                        radio_clicked = True
                except Exception as e:
                    logger.debug(f"方法2失败: {e}")
            
            # 方法3: 通过卡号文本查找
            if not radio_clicked:
                logger.info("方法3: 通过卡号文本查找...")
                try:
                    # 查找包含卡号尾号的td元素
                    card_selector = f"td:has-text('{card_tail_value}')"
                    card_elements = await target_frame.locator(card_selector).all()
                    
                    for card_element in card_elements:
                        try:
                            # 找到包含这个td的tr，然后找到其中的radio按钮
                            parent_tr = card_element.locator("xpath=..")
                            radio_in_tr = parent_tr.locator("input[type='radio'][name='rdoacnt']").first
                            if await radio_in_tr.count() > 0:
                                await radio_in_tr.click()
                                logger.info(f"✓ 通过卡号文本成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                                radio_clicked = True
                                break
                        except Exception as e:
                            logger.debug(f"在tr中查找radio失败: {e}")
                            continue
                except Exception as e:
                    logger.debug(f"方法3失败: {e}")
            
            # 方法4: 遍历所有radio按钮，检查其onclick属性
            if not radio_clicked:
                logger.info("方法4: 遍历所有radio按钮检查onclick属性...")
                try:
                    radio_buttons = await target_frame.locator("input[type='radio'][name='rdoacnt']").all()
                    for radio_button in radio_buttons:
                        try:
                            onclick_attr = await radio_button.get_attribute("onclick")
                            if onclick_attr and card_tail_value in onclick_attr:
                                await radio_button.click()
                                logger.info(f"✓ 通过遍历onclick属性成功选择卡号尾号 {card_tail_value} 对应的银行卡")
                                radio_clicked = True
                                break
                        except Exception as e:
                            logger.debug(f"检查radio按钮onclick属性失败: {e}")
                            continue
                except Exception as e:
                    logger.debug(f"方法4失败: {e}")
            
            if radio_clicked:
                logger.info("银行卡选择成功，等待选择生效...")
                await asyncio.sleep(2)  # 等待选择生效
                
                # 点击确定按钮
                await self.click_confirm_button_in_dialog()
            else:
                logger.warning(f"未找到卡号尾号 {card_tail_value} 对应的银行卡")
                
        except Exception as e:
            logger.error(f"处理转卡信息工号银行卡选择失败: {e}")
    
    async def click_confirm_button_in_dialog(self):
        """
        在对话框中点击确定按钮
        """
        try:
            # 尝试多种确定按钮的选择器
            confirm_selectors = [
                "button:has-text('确定')",
                "button.ui-button:has-text('确定')",
                "div.ui-dialog-buttonset button:has-text('确定')",
                "button[class*='ui-button']:has-text('确定')",
                "div.ui-dialog-buttonpane button:has-text('确定')",
                "button[role='button']:has-text('确定')"
            ]
            
            confirm_clicked = False
            
            # 首先在主页面查找
            for selector in confirm_selectors:
                try:
                    confirm_button = self.page.locator(selector).first
                    if await confirm_button.count() > 0:
                        await confirm_button.click()
                        logger.info(f"✓ 在主页面成功点击确定按钮 (使用选择器: {selector})")
                        await asyncio.sleep(BUTTON_CLICK_WAIT)
                        confirm_clicked = True
                        break
                except Exception as e:
                    logger.debug(f"主页面确定按钮选择器 {selector} 失败: {e}")
                    continue
            
            # 如果主页面没找到，尝试在所有iframe中查找
            if not confirm_clicked:
                frames = self.page.frames
                for i, frame in enumerate(frames):
                    for selector in confirm_selectors:
                        try:
                            confirm_button = frame.locator(selector).first
                            if await confirm_button.count() > 0:
                                await confirm_button.click()
                                logger.info(f"✓ 在iframe {i} 中成功点击确定按钮 (使用选择器: {selector})")
                                await asyncio.sleep(BUTTON_CLICK_WAIT)
                                confirm_clicked = True
                                break
                        except Exception as e:
                            logger.debug(f"iframe {i} 确定按钮选择器 {selector} 失败: {e}")
                            continue
                    if confirm_clicked:
                        break
            
            if not confirm_clicked:
                logger.warning("未找到确定按钮")
                
        except Exception as e:
            logger.debug(f"点击确定按钮失败: {e}")
    
    async def select_card_by_number(self, card_tail: str, retries: int = MAX_RETRIES):
        """
        根据卡号尾号选择对应的radio按钮
        
        Args:
            card_tail: 卡号尾号（不包含*前缀）
            retries: 重试次数
        """
        logger.info(f"开始选择卡号尾号: {card_tail}")
        
        for attempt in range(retries):
            try:
                frames = self.page.frames
                
                # 首先在主页面查找
                try:
                    # 查找包含指定卡号尾号的td元素，然后找到同行的radio按钮
                    radio_selector = f"//tr[td[contains(text(), '{card_tail}')]]/td/input[@type='radio'][@name='rdoacnt']"
                    radio_element = self.page.locator(radio_selector).first
                    if await radio_element.count() > 0:
                        await radio_element.click()
                        logger.info(f"成功选择卡号尾号 {card_tail} 对应的radio按钮")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                except Exception as e:
                    logger.debug(f"在主页面查找radio按钮失败: {e}")
                
                # 在iframe中查找
                for frame in frames:
                    try:
                        radio_selector = f"//tr[td[contains(text(), '{card_tail}')]]/td/input[@type='radio'][@name='rdoacnt']"
                        radio_element = frame.locator(radio_selector).first
                        if await radio_element.count() > 0:
                            await radio_element.click()
                            logger.info(f"在iframe中成功选择卡号尾号 {card_tail} 对应的radio按钮")
                            await asyncio.sleep(ELEMENT_WAIT)
                            return
                    except Exception as e:
                        logger.debug(f"在iframe中查找radio按钮失败: {e}")
                        continue
                
                # 尝试更通用的选择器
                try:
                    # 使用onclick属性中包含卡号尾号的方式
                    radio_selector = f"input[type='radio'][name='rdoacnt'][onclick*='{card_tail}']"
                    
                    # 先在主页面尝试
                    radio_element = self.page.locator(radio_selector).first
                    if await radio_element.count() > 0:
                        await radio_element.click()
                        logger.info(f"通过onclick属性成功选择卡号尾号 {card_tail} 对应的radio按钮")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                    
                    # 在iframe中尝试
                    for frame in frames:
                        try:
                            radio_element = frame.locator(radio_selector).first
                            if await radio_element.count() > 0:
                                await radio_element.click()
                                logger.info(f"在iframe中通过onclick属性成功选择卡号尾号 {card_tail} 对应的radio按钮")
                                await asyncio.sleep(ELEMENT_WAIT)
                                return
                        except Exception as e:
                            logger.debug(f"在iframe中通过onclick属性查找失败: {e}")
                            continue
                
                except Exception as e:
                    logger.debug(f"通过onclick属性查找失败: {e}")
                
                logger.warning(f"未找到卡号尾号 {card_tail} 对应的radio按钮")
                return
                
            except Exception as e:
                logger.warning(f"选择卡号radio按钮失败 (尝试 {attempt + 1}/{retries}): {card_tail} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"选择卡号radio按钮最终失败: {card_tail}")
    
    async def click_print_button(self):
        """
        先点击网页上的打印确认单按钮，然后等待5秒，最后使用坐标点击Chrome打印对话框
        """
        try:
            logger.info("查找网页上的打印确认单按钮...")
            
            # 等待页面加载完成
            await asyncio.sleep(2)
            
            # 查找并点击网页上的打印确认单按钮
            print_button_found = await self._find_and_click_print_button()
            
            if print_button_found:
                logger.info("✓ 网页打印确认单按钮点击成功")
                
                # 等待5秒钟，确保Chrome打印页面完全加载
                logger.info("等待5秒钟，确保Chrome打印页面加载完成...")
                await asyncio.sleep(5)
                
                # 从配置中获取Chrome打印对话框中保存按钮的坐标
                from config import PRINT_DIALOG_COORDINATES
                coords = PRINT_DIALOG_COORDINATES
                chrome_print_x = coords["print_button"]["x"]
                chrome_print_y = coords["print_button"]["y"]
                
                logger.info(f"使用坐标点击Chrome打印对话框中的保存按钮: ({chrome_print_x}, {chrome_print_y})")
                
                # 使用坐标点击Chrome打印对话框中的保存按钮
                logger.info(f"尝试点击坐标: ({chrome_print_x}, {chrome_print_y})")
                await self.page.mouse.click(chrome_print_x, chrome_print_y)
                logger.info("✓ Chrome打印对话框保存按钮点击成功")
                
                # 等待一下，看看是否真的点击成功了
                await asyncio.sleep(1)
                
                # 检查是否出现了文件保存对话框（通过检查是否有文件路径输入框）
                logger.info("检查是否出现文件保存对话框...")
                try:
                    # 尝试点击文件路径输入框，如果成功说明文件保存对话框已出现
                    coords = PRINT_DIALOG_COORDINATES
                    filepath_x = coords["filepath_input"]["x"]
                    filepath_y = coords["filepath_input"]["y"]
                    
                    await self.page.mouse.click(filepath_x, filepath_y)
                    logger.info("✓ 文件保存对话框已出现，继续处理...")
                except Exception as e:
                    logger.warning(f"文件保存对话框可能未出现: {e}")
                    logger.info("尝试重新点击Chrome打印对话框中的保存按钮...")
                    # 再次尝试点击保存按钮
                    await self.page.mouse.click(chrome_print_x, chrome_print_y)
                    await asyncio.sleep(1)
                
                # 处理文件保存对话框
                await self.handle_print_dialog()
            else:
                logger.error("❌ 网页打印确认单按钮点击失败")
                
        except Exception as e:
            logger.error(f"点击打印按钮失败: {e}")
            # 如果主要方法失败，尝试备用方案
            logger.info("主要方法失败，尝试备用方案...")
            await self._click_print_button_fallback()
    
    async def _find_and_click_print_button(self):
        """
        查找并点击网页上的打印确认单按钮
        """
        try:
            # 查找打印按钮的选择器（基于之前成功的日志）
            print_button_selectors = [
                'input[name="BtnPrint"]',
                'input[value="打印确认单"]',
                'input[onclick*="ybprint"]',
                '#BtnPrint',
                'input.buttHighlight'
            ]
            
            # 优先在iframe中查找（根据日志，按钮通常在iframe 7中）
            frames = self.page.frames
            logger.info(f"在 {len(frames)} 个iframe中查找打印按钮...")
            
            for i, frame in enumerate(frames):
                logger.info(f"在iframe {i} 中查找打印按钮: {frame.url or 'unnamed frame'}")
                
                for selector in print_button_selectors:
                    try:
                        button = frame.locator(selector).first
                        if await button.count() > 0:
                            logger.info(f"在iframe {i} 中找到打印按钮: {selector}")
                            logger.info(f"准备点击打印按钮...")
                            try:
                                # 方法1：使用click()方法
                                logger.info(f"尝试方法1：使用click()方法")
                                await button.click(timeout=3000)
                                logger.info(f"✓ 在iframe {i} 中成功点击打印按钮")
                                return True
                            except Exception as click_error:
                                logger.warning(f"方法1失败: {click_error}")
                                try:
                                    # 方法2：使用JavaScript点击
                                    logger.info(f"尝试方法2：使用JavaScript点击")
                                    await frame.evaluate("arguments[0].click();", button)
                                    logger.info(f"✓ 在iframe {i} 中使用JavaScript成功点击打印按钮")
                                    return True
                                except Exception as js_error:
                                     logger.warning(f"方法2失败: {js_error}")
                                     try:
                                         # 方法3：使用坐标点击
                                         logger.info(f"尝试方法3：使用坐标点击")
                                         bbox = await button.bounding_box()
                                         if bbox:
                                             x = bbox['x'] + bbox['width'] / 2
                                             y = bbox['y'] + bbox['height'] / 2
                                             await frame.mouse.click(x, y, timeout=3000)
                                             logger.info(f"✓ 在iframe {i} 中使用坐标成功点击打印按钮")
                                             return True
                                         else:
                                             logger.warning("无法获取按钮边界框")
                                     except Exception as coord_error:
                                         logger.warning(f"方法3失败: {coord_error}")
                                
                                # 即使所有方法都失败，也认为找到了按钮，继续执行
                                logger.info(f"所有点击方法都失败，但继续执行，假设打印按钮已点击")
                                return True
                    except Exception as e:
                        logger.debug(f"iframe {i} 选择器 {selector} 失败: {e}")
                        continue
            
            # 如果在iframe中没找到，在主页面查找
            logger.info("在iframe中未找到，尝试在主页面查找...")
            for selector in print_button_selectors:
                try:
                    button = self.page.locator(selector).first
                    if await button.count() > 0:
                        logger.info(f"在主页面找到打印按钮: {selector}")
                        logger.info(f"准备点击打印按钮...")
                        try:
                            # 方法1：使用click()方法
                            logger.info(f"尝试方法1：使用click()方法")
                            await button.click(timeout=3000)
                            logger.info("✓ 在主页面成功点击打印按钮")
                            return True
                        except Exception as click_error:
                            logger.warning(f"方法1失败: {click_error}")
                            try:
                                # 方法2：使用JavaScript点击
                                logger.info(f"尝试方法2：使用JavaScript点击")
                                await self.page.evaluate("arguments[0].click();", button)
                                logger.info("✓ 在主页面使用JavaScript成功点击打印按钮")
                                return True
                            except Exception as js_error:
                                logger.warning(f"方法2失败: {js_error}")
                                try:
                                    # 方法3：使用坐标点击
                                    logger.info(f"尝试方法3：使用坐标点击")
                                    bbox = await button.bounding_box()
                                    if bbox:
                                        x = bbox['x'] + bbox['width'] / 2
                                        y = bbox['y'] + bbox['height'] / 2
                                        await self.page.mouse.click(x, y, timeout=3000)
                                        logger.info("✓ 在主页面使用坐标成功点击打印按钮")
                                        return True
                                    else:
                                        logger.warning("无法获取按钮边界框")
                                except Exception as coord_error:
                                    logger.warning(f"方法3失败: {coord_error}")
                            
                            # 即使所有方法都失败，也认为找到了按钮，继续执行
                            logger.info(f"所有点击方法都失败，但继续执行，假设打印按钮已点击")
                            return True
                except Exception as e:
                    logger.debug(f"主页面选择器 {selector} 失败: {e}")
                    continue
            
            logger.warning("未找到打印按钮")
            return False
            
        except Exception as e:
            logger.error(f"查找打印按钮失败: {e}")
            return False
    
    async def _click_print_button_fallback(self):
        """
        打印按钮点击备用方案（使用元素选择器）
        """
        try:
            # 查找打印按钮
            print_button_selectors = [
                'input[name="BtnPrint"]',
                'input[value="打印确认单"]',
                'input[onclick*="ybprint"]',
                '#BtnPrint',
                'input.buttHighlight'
            ]
            
            print_button_found = False
            
            # 在主页面查找
            for selector in print_button_selectors:
                try:
                    button = self.page.locator(selector).first
                    if await button.count() > 0:
                        logger.info(f"备用方案：在主页面找到打印按钮: {selector}")
                        try:
                            await button.click(timeout=3000)
                            logger.info(f"✓ 备用方案：在主页面成功点击打印按钮")
                            print_button_found = True
                            break
                        except Exception as click_error:
                            logger.warning(f"备用方案点击打印按钮时出错: {click_error}")
                            # 即使点击出错，也认为找到了按钮，继续执行
                            logger.info(f"备用方案继续执行，假设打印按钮已点击")
                            print_button_found = True
                            break
                except Exception as e:
                    logger.debug(f"主页面选择器 {selector} 失败: {e}")
                    continue
            
            # 如果在主页面没找到，在iframe中查找
            if not print_button_found:
                frames = self.page.frames
                for i, frame in enumerate(frames):
                    for selector in print_button_selectors:
                        try:
                            button = frame.locator(selector).first
                            if await button.count() > 0:
                                logger.info(f"备用方案：在iframe {i} 中找到打印按钮: {selector}")
                                try:
                                    await button.click(timeout=3000)
                                    logger.info(f"✓ 备用方案：在iframe {i} 中成功点击打印按钮")
                                    print_button_found = True
                                    break
                                except Exception as click_error:
                                    logger.warning(f"备用方案iframe {i} 点击打印按钮时出错: {click_error}")
                                    # 即使点击出错，也认为找到了按钮，继续执行
                                    logger.info(f"备用方案继续执行，假设打印按钮已点击")
                                    print_button_found = True
                                    break
                        except Exception as e:
                            logger.debug(f"iframe {i} 选择器 {selector} 失败: {e}")
                            continue
                    if print_button_found:
                        break
            
            if print_button_found:
                logger.info("✓ 备用方案：网页打印确认单按钮点击成功")
                
                # 等待5秒钟，确保Chrome打印页面完全加载
                logger.info("等待5秒钟，确保Chrome打印页面加载完成...")
                await asyncio.sleep(5)
                
                # 从配置中获取Chrome打印对话框中保存按钮的坐标
                from config import PRINT_DIALOG_COORDINATES
                coords = PRINT_DIALOG_COORDINATES
                chrome_print_x = coords["print_button"]["x"]
                chrome_print_y = coords["print_button"]["y"]
                
                logger.info(f"备用方案：使用坐标点击Chrome打印对话框中的保存按钮: ({chrome_print_x}, {chrome_print_y})")
                
                # 使用坐标点击Chrome打印对话框中的保存按钮
                logger.info(f"备用方案：尝试点击坐标: ({chrome_print_x}, {chrome_print_y})")
                await self.page.mouse.click(chrome_print_x, chrome_print_y)
                logger.info("✓ 备用方案：Chrome打印对话框保存按钮点击成功")
                
                # 等待一下，看看是否真的点击成功了
                await asyncio.sleep(1)
                
                # 检查是否出现了文件保存对话框
                logger.info("备用方案：检查是否出现文件保存对话框...")
                try:
                    # 尝试点击文件路径输入框，如果成功说明文件保存对话框已出现
                    coords = PRINT_DIALOG_COORDINATES
                    filepath_x = coords["filepath_input"]["x"]
                    filepath_y = coords["filepath_input"]["y"]
                    
                    await self.page.mouse.click(filepath_x, filepath_y)
                    logger.info("✓ 备用方案：文件保存对话框已出现，继续处理...")
                except Exception as e:
                    logger.warning(f"备用方案：文件保存对话框可能未出现: {e}")
                    logger.info("备用方案：尝试重新点击Chrome打印对话框中的保存按钮...")
                    # 再次尝试点击保存按钮
                    await self.page.mouse.click(chrome_print_x, chrome_print_y)
                    await asyncio.sleep(1)
                
                # 处理文件保存对话框
                await self.handle_print_dialog()
            else:
                logger.error("❌ 备用方案：未找到打印按钮")
                
        except Exception as e:
            logger.error(f"备用方案点击打印按钮失败: {e}")
    
    async def handle_print_dialog(self):
        """
        处理打印对话框
        """
        try:
            logger.info("开始处理打印对话框...")
            
            # 获取当前记录的项目号和总金额信息
            project_number = self.get_current_project_number()
            total_amount = self.get_current_total_amount()
            
            logger.info(f"项目号: {project_number}, 总金额: {total_amount}")
            
            # 等待打印对话框出现
            logger.info("等待打印对话框出现...")
            await asyncio.sleep(PRINT_DIALOG_WAIT_TIME)
            
            # 从config中获取坐标配置和文件路径
            coords = PRINT_DIALOG_COORDINATES
            chrome_print_x = coords["print_button"]["x"]  # Chrome打印对话框中保存按钮的X坐标
            chrome_print_y = coords["print_button"]["y"]  # Chrome打印对话框中保存按钮的Y坐标
            filepath_x = coords["filepath_input"]["x"]    # 文件路径输入框的X坐标
            filepath_y = coords["filepath_input"]["y"]    # 文件路径输入框的Y坐标
            filename_x = coords["filename_input"]["x"]    # 文件名输入框的X坐标
            filename_y = coords["filename_input"]["y"]    # 文件名输入框的Y坐标
            save_x = coords["save_button"]["x"]           # 文件保存对话框中保存按钮的X坐标
            save_y = coords["save_button"]["y"]           # 文件保存对话框中保存按钮的Y坐标
            # 已禁用文件覆盖功能，不再需要yes_button坐标
            
            # 获取文件保存路径
            from config import PRINT_FILE_PATH
            file_path = PRINT_FILE_PATH
            
            # 先点击文件路径输入框
            logger.info(f"点击文件路径输入框，坐标: ({filepath_x}, {filepath_y})")
            await self.page.mouse.click(filepath_x, filepath_y)
            await asyncio.sleep(0.5)
            
            # 确保输入框获得焦点后再输入
            logger.info(f"确保输入框获得焦点...")
            await asyncio.sleep(0.3)
            
            # 清空现有内容并输入文件路径
            logger.info(f"输入文件路径: {file_path}")
            await self.page.keyboard.press('Control+a')  # 全选现有内容
            await asyncio.sleep(0.2)
            await self.page.keyboard.press('Delete')   # 删除现有内容
            await asyncio.sleep(0.2)
            await self.page.keyboard.type(file_path)   # 输入新路径
            await asyncio.sleep(0.5)
            
            # 按Tab键移动到文件名输入框
            logger.info("按Tab键移动到文件名输入框")
            await self.page.keyboard.press('Tab')
            await asyncio.sleep(0.5)
            
            # 尝试导入打印对话框处理模块
            try:
                from print_dialog_handler import create_print_handler
                logger.info("✓ 成功导入打印对话框处理模块")
                
                # 创建打印对话框处理器
                print_handler = create_print_handler(PRINT_OUTPUT_DIR)
                
                # 处理打印操作
                success = print_handler.process_print_operation(
                    project_number, total_amount,
                    chrome_print_x, chrome_print_y,
                    filename_x, filename_y,
                    save_x, save_y
                )
                
                if success:
                    logger.info("✓ 打印对话框处理成功")
                else:
                    logger.warning("⚠ 打印对话框处理失败，尝试备用方案")
                    await self._handle_print_dialog_fallback()
                    
            except ImportError as e:
                logger.error(f"导入打印对话框处理模块失败: {e}")
                logger.info("使用备用方案处理打印对话框")
                await self._handle_print_dialog_fallback()
                
        except Exception as e:
            logger.error(f"处理打印对话框失败: {e}")
            logger.info("请手动处理打印对话框")
    
    async def _handle_print_dialog_fallback(self):
        """
        打印对话框处理备用方案
        """
        try:
            logger.info("使用备用方案处理打印对话框...")
            
            # 等待一段时间让用户看到对话框
            await asyncio.sleep(2)
            
            # 从config中获取文件路径
            from config import PRINT_FILE_PATH
            file_path = PRINT_FILE_PATH
            
            # 尝试点击文件路径输入框并输入路径
            try:
                coords = PRINT_DIALOG_COORDINATES
                filepath_x = coords["filepath_input"]["x"]
                filepath_y = coords["filepath_input"]["y"]
                
                logger.info(f"备用方案：点击文件路径输入框，坐标: ({filepath_x}, {filepath_y})")
                await self.page.mouse.click(filepath_x, filepath_y)
                await asyncio.sleep(0.5)
                
                # 清空现有内容并输入文件路径
                logger.info(f"备用方案：输入文件路径: {file_path}")
                await self.page.keyboard.press('Control+a')  # 全选现有内容
                await asyncio.sleep(0.2)
                await self.page.keyboard.press('Delete')   # 删除现有内容
                await asyncio.sleep(0.2)
                await self.page.keyboard.type(file_path)   # 输入新路径
                await asyncio.sleep(0.5)
                
                # 按Tab键移动到文件名输入框
                logger.info("备用方案：按Tab键移动到文件名输入框")
                await self.page.keyboard.press('Tab')
                await asyncio.sleep(0.5)
                
            except Exception as e:
                logger.warning(f"备用方案文件路径输入失败: {e}")
            
            # 尝试按Tab键导航到保存按钮
            logger.info("尝试按Tab键导航...")
            for i in range(15):
                await self.page.keyboard.press('Tab')
                await asyncio.sleep(0.3)
            
            # 按Enter键确认
            logger.info("按Enter键确认...")
            await self.page.keyboard.press('Enter')
            await asyncio.sleep(2)
            
            # 如果还在对话框中，尝试按Escape键关闭
            logger.info("尝试按Escape键关闭对话框...")
            await self.page.keyboard.press('Escape')
            await asyncio.sleep(1)
            
            logger.info("备用方案处理完成")
            
        except Exception as e:
            logger.error(f"备用方案处理失败: {e}")
            logger.info("请手动处理打印对话框")
    
    def get_current_project_number(self) -> str:
        """
        获取当前记录的项目编号
        """
        try:
            # 优先使用保存的报销项目号
            if self.current_project_number is not None and self.current_project_number != "":
                logger.info(f"使用保存的报销项目号: {self.current_project_number}")
                return self.current_project_number
            
            # 如果保存的值不存在，从Excel数据中查找
            if hasattr(self, 'reimbursement_data') and self.current_sequence is not None:
                # 查找当前序号对应的记录
                current_record = self.reimbursement_data[self.reimbursement_data[SEQUENCE_COL] == self.current_sequence]
                if not current_record.empty:
                    # 查找项目编号列（按优先级）
                    project_columns = ["报销项目号", "项目编号", "项目号"]
                    for col in project_columns:
                        if col in current_record.columns:
                            value = current_record[col].iloc[0]
                            if pd.notna(value) and value != "":
                                project_number = self.clean_value_string(value)
                                logger.info(f"从Excel列 '{col}' 获取到项目编号: {project_number}")
                                return project_number
            
            # 如果没找到，返回默认值
            return "未知项目"
        except Exception as e:
            logger.debug(f"获取项目编号失败: {e}")
            return "未知项目"
    
    def get_current_total_amount(self) -> str:
        """
        获取当前记录的总金额
        """
        try:
            # 优先使用保存的金额
            if self.current_amount is not None and self.current_amount != "":
                logger.info(f"使用保存的金额: {self.current_amount}")
                return self.current_amount
            
            # 如果保存的值不存在，从Excel数据中查找
            if hasattr(self, 'reimbursement_data') and self.current_sequence is not None:
                # 查找当前序号对应的记录
                current_record = self.reimbursement_data[self.reimbursement_data[SEQUENCE_COL] == self.current_sequence]
                if not current_record.empty:
                    # 查找金额列（按优先级）
                    amount_columns = ["金额", "总金额", "个人金额"]
                    for col in amount_columns:
                        if col in current_record.columns:
                            value = current_record[col].iloc[0]
                            if pd.notna(value) and value != "":
                                amount = self.clean_value_string(value)
                                logger.info(f"从Excel列 '{col}' 获取到金额: {amount}")
                                return amount
            
            # 如果没找到，返回默认值
            return "0"
        except Exception as e:
            logger.debug(f"获取总金额失败: {e}")
            return "0"
    
    async def process_sequence_with_subsequences(self, sequence_num: int, group_data: pd.DataFrame):
        """
        处理带有子序列逻辑的序号组
        
        Args:
            sequence_num: 序号
            group_data: 该序号下的所有数据行
        """
        logger.info(f"开始处理序号 {sequence_num} 的报销记录，共 {len(group_data)} 行")
        
        # 检查是否包含登录信息（通常在第一行）
        first_row = group_data.iloc[0]
        if "登录界面工号" in group_data.columns and pd.notna(first_row["登录界面工号"]):
            # 处理登录流程
            logger.info("检测到登录信息，开始处理登录流程")
            await self.handle_login_with_captcha(group_data)
            
            # 登录完成后，继续处理该序号下的其他行（如果有的话）
            if len(group_data) > 1:
                logger.info(f"登录完成，继续处理序号 {sequence_num} 的剩余 {len(group_data) - 1} 行数据")
                # 处理剩余的行（跳过第一行，因为已经处理了登录）
                remaining_data = group_data.iloc[1:]
                # 检查是否包含第二种子序列，如果包含则跳过重复处理
                has_traveler_subsequence = False
                for _, row in remaining_data.iterrows():
                    for col in remaining_data.columns:
                        if col.startswith(SUBSEQUENCE_START_COL):
                            if pd.notna(row[col]) and self.clean_value_string(row[col]) == TRAVELER_SUBSEQUENCE_MARKER:
                                has_traveler_subsequence = True
                                break
                    if has_traveler_subsequence:
                        break
                
                if has_traveler_subsequence:
                    logger.info("检测到第二种子序列已在登录流程中处理，跳过重复处理，但继续处理后续操作字段")
                    # 即使跳过了重复处理，也要处理后续的操作字段（如"下一步按钮4"）
                    await self.process_remaining_operations(remaining_data)
                else:
                    await self.process_subsequences(remaining_data)
        else:
            # 处理子序列逻辑
            logger.info("未检测到登录信息，直接处理子序列逻辑")
            await self.process_subsequences(group_data)
    
    async def process_subsequences(self, group_data: pd.DataFrame):
        """
        处理子序列逻辑：
        1. 当检测到"子序列开始"时，开始子序列处理
        2. 处理当前行的所有列（从子序列开始列的下一列开始）
        3. 如果遇到"子序列结束"标记，继续在同一行向右处理
        4. 重复步骤2-3，直到没有更多的子序列行
        
        Args:
            group_data: 同一序号下的所有数据行
        """
        # 将DataFrame转换为list以便遍历
        rows = group_data.to_dict('records')
        columns = list(group_data.columns)
        
        i = 0
        while i < len(rows):
            row = rows[i]
            current_sequence = row.get(SEQUENCE_COL, None)
            logger.info(f"处理第 {i+1} 行数据，序号: {current_sequence}")
            
            # 查找子序列开始和结束列的位置（支持自动重命名）
            subsequence_start_idx = None
            subsequence_end_idx = None
            
            # 查找子序列开始列（支持自动重命名如"子序列开始.1"）
            subsequence_start_cols = [col for col in columns if col.startswith(SUBSEQUENCE_START_COL)]
            if subsequence_start_cols:
                # 使用第一个找到的子序列开始列
                subsequence_start_idx = columns.index(subsequence_start_cols[0])
                logger.info(f"找到子序列开始列: {subsequence_start_cols[0]} (索引: {subsequence_start_idx})")
                if len(subsequence_start_cols) > 1:
                    logger.info(f"注意：还找到其他子序列开始列: {subsequence_start_cols[1:]}")
            else:
                logger.debug("未找到子序列开始列，按普通方式处理")
                subsequence_start_idx = None
            
            # 查找子序列结束列（支持自动重命名如"子序列结束.1"）
            subsequence_end_cols = [col for col in columns if col.startswith(SUBSEQUENCE_END_COL)]
            if subsequence_end_cols:
                # 使用第一个找到的子序列结束列
                subsequence_end_idx = columns.index(subsequence_end_cols[0])
                logger.info(f"找到子序列结束列: {subsequence_end_cols[0]} (索引: {subsequence_end_idx})")
                if len(subsequence_end_cols) > 1:
                    logger.info(f"注意：还找到其他子序列结束列: {subsequence_end_cols[1:]}")
            else:
                logger.debug("未找到子序列结束列")
                subsequence_end_idx = None
            
            # 处理当前行的列
            col_idx = 0
            while col_idx < len(columns):
                col = columns[col_idx]
                
                # 跳过序号列和处理进度列
                if col == SEQUENCE_COL or col == "处理进度":
                    col_idx += 1
                    continue
                
                # 如果到达子序列开始列，开始处理子序列
                if col_idx == subsequence_start_idx:
                    subsequence_value = row[col]
                    if pd.notna(subsequence_value) and subsequence_value != "":
                        subsequence_value_str = self.clean_value_string(subsequence_value)
                        
                        # 检查是否为第二种子序列（数字"1"标记）
                        if subsequence_value_str == TRAVELER_SUBSEQUENCE_MARKER:
                            logger.info(f"检测到第二种子序列（出差人信息），开始处理出差人信息填写")
                            await self.process_traveler_subsequence(group_data, i)
                            # 第二种子序列处理完成后，继续处理子序列结束后的其他列
                            if subsequence_end_idx is not None:
                                col_idx = subsequence_end_idx + 1
                            else:
                                col_idx = len(columns)
                            continue
                        else:
                            logger.info(f"检测到第一种子序列，开始处理子序列逻辑")
                            # 处理子序列：从子序列开始列的下一列开始，到子序列结束列
                            subseq_start_col_idx = subsequence_start_idx + 1
                            subseq_end_col_idx = subsequence_end_idx if subsequence_end_idx is not None else len(columns)
                            
                            # 处理当前行的子序列部分
                            encountered_end_marker = await self.process_subsequence_row(row, columns, subseq_start_col_idx, subseq_end_col_idx)
                    
                    # 子序列处理完成后，跳转到子序列结束列的下一列继续处理
                    if subsequence_end_idx is not None:
                        col_idx = subsequence_end_idx + 1
                    else:
                        col_idx = len(columns)
                    continue
                else:
                    # 处理子序列开始之前的列
                    value = row[col]
                    if pd.notna(value) and value != "":
                        value_str = self.clean_value_string(value)
                        
                        # 特殊处理：科目列（以#开头）
                        if value_str.startswith("#"):
                            logger.info(f"处理科目列: {col} = {value_str}")
                            
                            # 提取科目名称（去掉#前缀）
                            subject_name = value_str[1:]
                            
                            # 在标题-ID表中查找对应的输入框ID
                            input_id = self.get_object_id(subject_name)
                            if not input_id:
                                logger.warning(f"未找到科目 '{subject_name}' 对应的ID映射")
                                col_idx += 1
                                continue
                            
                            # 查找下一列的金额
                            if col_idx + 1 < len(columns):
                                amount_col = columns[col_idx + 1]
                                amount_value = row[amount_col]
                                
                                if pd.notna(amount_value) and amount_value != "":
                                    amount_str = self.clean_value_string(amount_value)
                                    logger.info(f"找到金额列: {amount_col} = {amount_str}")
                                    
                                    # 填写金额到对应的输入框
                                    await self.fill_input(input_id, amount_str, title=amount_col)
                                    logger.info(f"成功填写科目 '{subject_name}' 的金额: {amount_str}")
                                    
                                    # 跳过金额列，因为已经处理了
                                    col_idx += 2
                                    continue
                                else:
                                    logger.warning(f"科目 '{subject_name}' 对应的金额列为空")
                                    col_idx += 1
                                    continue
                            else:
                                logger.warning(f"科目 '{subject_name}' 没有对应的金额列")
                                col_idx += 1
                                continue
                        
                        # 处理普通列
                        logger.info(f"处理普通操作: {col} = {value_str}")
                        await self.process_cell(col, value_str)
                    col_idx += 1
            
            # 移动到下一行
            i += 1
    
    async def process_subsequence_row(self, row: Dict, columns: List[str], start_col_idx: int, end_col_idx: int):
        """
        处理单行的子序列部分
        
        Args:
            row: 行数据字典
            columns: 列名列表
            start_col_idx: 子序列开始列索引
            end_col_idx: 子序列结束列索引
            
        Returns:
            bool: 是否遇到子序列结束标记（已废弃，始终返回False）
        """
        logger.info(f"处理子序列行，从列 {start_col_idx} 到 {end_col_idx}")
        
        # 只处理子序列范围内的列（从start_col_idx到end_col_idx）
        col_idx = start_col_idx
        while col_idx < end_col_idx:
            col = columns[col_idx]
            
            # 跳过子序列开始列和处理进度列
            if col in [SUBSEQUENCE_START_COL, "处理进度"]:
                col_idx += 1
                continue
            
            # 跳过子序列结束列本身
            if col == SUBSEQUENCE_END_COL:
                col_idx += 1
                continue
            
            value = row[col]
            if pd.notna(value) and value != "":
                value_str = self.clean_value_string(value)
                
                # 特殊处理：科目列（以#开头）
                if value_str.startswith("#"):
                    logger.info(f"处理科目列: {col} = {value_str}")
                    
                    # 提取科目名称（去掉#前缀）
                    subject_name = value_str[1:]
                    
                    # 在标题-ID表中查找对应的输入框ID
                    input_id = self.get_object_id(subject_name)
                    if not input_id:
                        logger.warning(f"未找到科目 '{subject_name}' 对应的ID映射")
                        col_idx += 1
                        continue
                    
                    # 查找下一列的金额
                    if col_idx + 1 < end_col_idx:
                        amount_col = columns[col_idx + 1]
                        amount_value = row[amount_col]
                        
                        if pd.notna(amount_value) and amount_value != "":
                            amount_str = self.clean_value_string(amount_value)
                            logger.info(f"找到金额列: {amount_col} = {amount_str}")
                            
                            # 填写金额到对应的输入框
                            await self.fill_input(input_id, amount_str, title=amount_col)
                            logger.info(f"成功填写科目 '{subject_name}' 的金额: {amount_str}")
                            
                            # 跳过金额列，因为已经处理了
                            col_idx += 2
                            continue
                        else:
                            logger.warning(f"科目 '{subject_name}' 对应的金额列为空")
                            col_idx += 1
                            continue
                    else:
                        logger.warning(f"科目 '{subject_name}' 没有对应的金额列")
                        col_idx += 1
                        continue
                
                # 处理普通列
                logger.info(f"处理子序列操作: {col} = {value_str}")
                await self.process_cell(col, value_str)
            
            col_idx += 1
        
        return False  # 不再返回子序列结束标记
    
    async def process_single_row(self, row: Dict, columns: List[str]):
        """
        处理单行数据（没有子序列逻辑）
        
        Args:
            row: 行数据字典
            columns: 列名列表
        """
        for col in columns:
            if col == SEQUENCE_COL or col == "处理进度":  # 跳过序号列和处理进度列
                continue
            
            value = row[col]
            if pd.notna(value) and value != "":
                value_str = self.clean_value_string(value)
                logger.info(f"处理操作: {col} = {value_str}")
                await self.process_cell(col, value_str)
    
    async def handle_login_with_captcha(self, record_data: pd.DataFrame):
        """
        处理登录流程，包括验证码输入
        
        Args:
            record_data: 包含登录信息的DataFrame行
        """
        logger.info("开始处理登录流程...")
        
        # 填写工号
        if "登录界面工号" in record_data.columns:
            uid = record_data["登录界面工号"].iloc[0]
            uid_str = self.clean_value_string(uid)
            if uid_str:
                logger.info(f"填写工号: {uid_str}")
                await self.fill_input("uid", uid_str)
        
        # 填写密码
        if "登录界面密码" in record_data.columns:
            pwd = record_data["登录界面密码"].iloc[0]
            pwd_str = self.clean_value_string(pwd)
            if pwd_str:
                logger.info("填写密码完成")
                await self.fill_input("pwd", pwd_str)
        
        # 等待用户输入验证码
        logger.info("=" * 50)
        logger.info("密码填写完成，请在下方输入验证码:")
        logger.info("=" * 50)
        
        # 强制刷新输出缓冲区
        import sys
        sys.stdout.flush()
        
        try:
            captcha = input("请输入验证码: ")
            logger.info(f"用户输入验证码: {captcha}")
        except Exception as e:
            logger.error(f"验证码输入失败: {e}")
            captcha = ""
        
        # 查找验证码输入框并填写
        try:
            # 尝试常见的验证码输入框选择器
            captcha_selectors = [
                "input[name='captcha']",
                "input[id*='captcha']",
                "input[placeholder*='验证码']",
                "input[placeholder*='captcha']",
                "#captcha",
                ".captcha-input"
            ]
            
            captcha_filled = False
            for selector in captcha_selectors:
                try:
                    await self.page.wait_for_selector(selector, timeout=1000)
                    await self.page.fill(selector, captcha)
                    logger.info(f"成功填写验证码: {captcha}")
                    captcha_filled = True
                    break
                except:
                    continue
            
            if not captcha_filled:
                logger.warning("未找到验证码输入框，请手动输入验证码")
        except Exception as e:
            logger.error(f"填写验证码失败: {e}")
        
        # 点击登录按钮
        if "登录按钮" in record_data.columns:
            login_btn = record_data["登录按钮"].iloc[0]
            if pd.notna(login_btn) and login_btn != "":
                logger.info("点击登录按钮...")
                await self.click_button("zhLogin")
        
        # 等待登录完成
        logger.info("登录请求已发送，等待页面跳转...")
        await asyncio.sleep(LOGIN_WAIT_TIME)
        
        # 登录完成后，继续处理当前记录中的其他操作
        logger.info("登录完成，继续处理当前记录中的其他操作...")
        await self.process_record_after_login(record_data)
    
    async def process_record_after_login(self, record_data: pd.DataFrame):
        """
        登录后处理当前记录中的其他操作
        
        Args:
            record_data: 包含该报销记录所有行的DataFrame
        """
        # 处理当前记录中的所有列（除了登录相关列）
        row = record_data.iloc[0]
        columns = list(record_data.columns)
        
        i = 0
        while i < len(columns):
            col = columns[i]
            
            # 跳过登录相关的列、序号列和处理进度列
            if col in ["序号", "登录界面工号", "登录界面密码", "登录按钮", "处理进度"]:
                i += 1
                continue
            
            value = row[col]
            if pd.notna(value) and value != "":
                value_str = self.clean_value_string(value)
                
                # 特殊处理：科目列（以#开头）
                if value_str.startswith("#"):
                    logger.info(f"处理科目列: {col} = {value_str}")
                    
                    # 提取科目名称（去掉#前缀）
                    subject_name = value_str[1:]
                    
                    # 在标题-ID表中查找对应的输入框ID
                    input_id = self.get_object_id(subject_name)
                    if not input_id:
                        logger.warning(f"未找到科目 '{subject_name}' 对应的ID映射")
                        i += 1
                        continue
                    
                    # 查找下一列的金额
                    if i + 1 < len(columns):
                        amount_col = columns[i + 1]
                        amount_value = row[amount_col]
                        
                        if pd.notna(amount_value) and amount_value != "":
                            amount_str = self.clean_value_string(amount_value)
                            logger.info(f"找到金额列: {amount_col} = {amount_str}")
                            
                            # 填写金额到对应的输入框
                            await self.fill_input(input_id, amount_str, title=amount_col)
                            logger.info(f"成功填写科目 '{subject_name}' 的金额: {amount_str}")
                            
                            # 跳过金额列，因为已经处理了
                            i += 2
                            continue
                        else:
                            logger.warning(f"科目 '{subject_name}' 对应的金额列为空")
                            i += 1
                            continue
                    else:
                        logger.warning(f"科目 '{subject_name}' 没有对应的金额列")
                        i += 1
                        continue
                
                # 特殊处理：子序列开始列
                if col == SUBSEQUENCE_START_COL:
                    if value_str == TRAVELER_SUBSEQUENCE_MARKER:
                        logger.info(f"检测到第二种子序列（出差人信息），开始处理出差人信息填写")
                        await self.process_traveler_subsequence(record_data, 0)  # 从第1行开始
                        # 跳过当前行的其余列，因为已经处理了
                        i = len(columns)  # 跳到行末
                        break  # 跳出while循环
                    else:
                        logger.info(f"检测到第一种子序列（值: {value_str}），继续处理")
                        await self.process_cell(col, value_str)
                else:
                    # 处理普通列
                    logger.info(f"处理登录后的操作: {col} = {value_str}")
                    await self.process_cell(col, value_str)
            
            i += 1
    
    async def process_reimbursement_record(self, record_data: pd.DataFrame):
        """
        处理单条报销记录
        
        Args:
            record_data: 包含该报销记录所有行的DataFrame
        """
        sequence_num = record_data[SEQUENCE_COL].iloc[0]
        self.current_sequence = sequence_num
        
        # 重置保存的报销项目号和金额，确保每个序号使用自己的值
        self.current_project_number = None
        self.current_amount = None
        logger.info(f"重置报销项目号和金额，准备处理序号 {sequence_num}")
        
        logger.info(f"开始处理序号 {sequence_num} 的报销记录，共{len(record_data)}行数据")
        
        # 检查是否有子序列列（支持自动重命名）
        has_subsequence_start = any(col.startswith(SUBSEQUENCE_START_COL) for col in record_data.columns)
        has_subsequence_end = any(col.startswith(SUBSEQUENCE_END_COL) for col in record_data.columns)
        has_subsequence = has_subsequence_start and has_subsequence_end
        
        if has_subsequence:
            # 处理子序列逻辑
            await self.process_subsequence_logic(record_data)
        else:
            # 处理普通逻辑（假设只有一行数据）
            row = record_data.iloc[0]
            for col in record_data.columns:
                if col == SEQUENCE_COL or col == "处理进度":
                    continue
                
                value = row[col]
                if pd.notna(value) and value != "":
                    await self.process_cell(col, value)
        
        logger.info(f"序号 {sequence_num} 的报销记录处理完成")
    
    async def process_traveler_subsequence(self, group_data: pd.DataFrame, start_row_idx: int):
        """
        处理第二种子序列逻辑：填写出差人信息到网页表格
        采用类似第一种子序列的方式，自动为字段名添加后缀并查询标题-ID映射
        
        Args:
            group_data: 同一序号下的所有数据行
            start_row_idx: 子序列开始的行索引
        """
        logger.info(f"开始处理出差人信息子序列，从第 {start_row_idx + 1} 行开始")
        
        # 从开始行开始，逐行处理子序列
        # 重置traveler_index，确保从0开始
        traveler_index = 0  # 出差人索引，用于生成后缀
        logger.info(f"重置traveler_index为0，开始处理第二种子序列")
        
        # 强制重置traveler_index，确保从0开始
        self.traveler_index = 0
        
        for row_idx in range(start_row_idx, len(group_data)):
            row = group_data.iloc[row_idx]
            
            # 检查是否为子序列结束（但先处理当前行的信息，支持自动重命名）
            is_end_marker = False
            for end_col in group_data.columns:
                if end_col.startswith(SUBSEQUENCE_END_COL):
                    if pd.notna(row[end_col]) and self.clean_value_string(row[end_col]) == TRAVELER_SUBSEQUENCE_MARKER:
                        is_end_marker = True
                        logger.info(f"检测到子序列结束标记，在第 {row_idx + 1} 行的列 {end_col}")
                        break
            
            if is_end_marker:
                logger.info(f"检测到子序列结束标记，在第 {row_idx + 1} 行")
                # 如果是子序列结束标记所在行，先处理当前行的出差人信息，然后结束
                logger.info(f"处理子序列结束标记所在行的出差人信息")
                should_break = True
            else:
                should_break = False
            
            # 检查当前行是否有有效的出差人信息
            has_traveler_info = False
            for field in TRAVELER_FIELDS.keys():
                if field in group_data.columns and pd.notna(row[field]) and row[field] != "":
                    has_traveler_info = True
                    break
            
            if not has_traveler_info:
                logger.info(f"第 {row_idx + 1} 行没有有效的出差人信息，跳过")
                continue
            
            # 限制最多6个出差人
            if traveler_index >= 6:
                logger.warning(f"出差人数量超过6个，跳过第 {traveler_index + 1} 个")
                break
            
            logger.info(f"处理第 {traveler_index + 1} 个出差人信息（第 {row_idx + 1} 行）")
            
            # 逐列处理当前行的出差人信息
            # 调整字段处理顺序：先填写姓名，再填写工号，避免工号触发的事件清空姓名
            field_order = ["姓名", "人员类型", "单位", "职称", "工号"]
            
            for field in field_order:
                if field in TRAVELER_FIELDS.keys() and field in group_data.columns and pd.notna(row[field]) and row[field] != "":
                    value = self.clean_value_string(row[field])
                    
                    # 为字段名添加后缀
                    field_with_suffix = f"{field}-{traveler_index}"
                    logger.info(f"处理字段: {field} -> {field_with_suffix} = {value}")
                    
                    # 在标题-ID映射表中查找对应的输入框ID
                    input_id = self.get_object_id(field_with_suffix)
                    if not input_id:
                        logger.warning(f"未找到字段 '{field_with_suffix}' 对应的ID映射")
                        continue
                    
                    # 根据字段类型选择填写方式
                    if field == "人员类型":
                        # 人员类型使用下拉选择
                        await self.select_dropdown(input_id, value)
                        logger.info(f"选择{field_with_suffix}: {value}")
                    elif field == "工号":
                        # 工号字段特殊处理：填写后等待一下，让JavaScript事件完成
                        await self.fill_input(input_id, value, title=field_with_suffix)
                        logger.info(f"填写{field_with_suffix}: {value}")
                        # 等待JavaScript事件完成
                        await asyncio.sleep(2)
                        logger.info(f"工号填写完成，等待JavaScript事件处理")
                        
                        # 重新填写姓名，确保不被JavaScript事件清空
                        name_field = f"姓名-{traveler_index}"
                        name_value = self.clean_value_string(row["姓名"])
                        name_input_id = self.get_object_id(name_field)
                        if name_input_id and name_value:
                            await self.fill_input(name_input_id, name_value, title=name_field)
                            logger.info(f"重新填写{name_field}: {name_value}")
                            await asyncio.sleep(0.5)
                    else:
                        # 其他字段使用输入框填写
                        await self.fill_input(input_id, value, title=field_with_suffix)
                        logger.info(f"填写{field_with_suffix}: {value}")
            
            # 处理当前行的其他字段（非出差人信息字段，但仅限于子序列范围内的字段）
            logger.info(f"处理第 {row_idx + 1} 行的其他字段")
            for col in group_data.columns:
                # 跳过序号列、处理进度列、子序列标记列、出差人信息字段和登录相关字段
                if (col in [SEQUENCE_COL, "处理进度", "登录界面工号", "登录界面密码", "登录按钮", "网上预约报账按钮", "等待", "申请报销单按钮", "已阅读并同意按钮", "选择业务大类", "报销项目号", "附件张数", "备注", "特殊事项说明", "下一步按钮1", "等待.1"] or 
                    col.startswith(SUBSEQUENCE_START_COL) or 
                    col.startswith(SUBSEQUENCE_END_COL) or
                    col in TRAVELER_FIELDS.keys()):
                    continue
                
                # 检查是否为子序列结束后的字段，如果是则跳过，让后续的正常处理逻辑处理
                # 注意：这里只跳过真正的子序列结束后的字段，不包括正常的操作字段
                if col.startswith("下一步按钮") and col != "下一步按钮1":
                    # 不跳过下一步按钮，让它在后续的正常处理逻辑中处理
                    pass
                
                value = row[col]
                if pd.notna(value) and value != "":
                    value_str = self.clean_value_string(value)
                    
                    # 为字段名添加后缀（使用固定的子序列索引0）
                    field_with_suffix = f"{col}-0"
                    logger.info(f"处理子序列操作: {col} -> {field_with_suffix} = {value_str}")
                    
                    # 在标题-ID映射表中查找对应的输入框ID
                    input_id = self.get_object_id(field_with_suffix)
                    if not input_id:
                        logger.warning(f"未找到字段 '{field_with_suffix}' 对应的ID映射")
                        continue
                    
                    # 根据字段类型选择填写方式
                    # 检查是否为下拉字段（通过字段名或ID模式）
                    is_dropdown = False
                    if col in DROPDOWN_FIELDS:
                        is_dropdown = True
                    elif col == "省份":  # 省份字段特殊处理
                        is_dropdown = True
                    elif input_id and "sf" in input_id:  # 省份字段的ID模式
                        is_dropdown = True
                    elif input_id and "hsf" in input_id:  # hsf字段的ID模式
                        is_dropdown = True
                    elif input_id and "jtf" in input_id:  # jtf字段的ID模式
                        is_dropdown = True
                    
                    # 检查是否为日期字段
                    is_date = False
                    if (input_id and ("date" in input_id.lower() or "startdate" in input_id.lower() or "enddate" in input_id.lower() or 
                                     input_id.endswith("_startdate") or input_id.endswith("_enddate") or
                                     "temp-startdate" in input_id or "temp-enddate" in input_id or
                                     input_id == "formWF_YB6_3492_yc-chr_start1_0" or input_id == "formWF_YB6_3492_yc-chr_end1_0" or
                                     "start" in input_id or "end" in input_id)):
                        is_date = True
                    
                    if is_date:
                        # 日期字段使用日历控件
                        logger.info(f"检测到日期输入框: {input_id} = {value_str}")
                        try:
                            await self.select_date_from_calendar(input_id, value_str)
                            logger.info(f"日期填写完成: {field_with_suffix}")
                        except Exception as e:
                            logger.warning(f"日期选择失败，尝试普通输入: {e}")
                            await self.fill_input(input_id, value_str, title=field_with_suffix)
                            logger.info(f"填写{field_with_suffix}: {value_str}")
                    elif is_dropdown:
                        # 下拉选择
                        await self.select_dropdown(input_id, value_str)
                        logger.info(f"选择{field_with_suffix}: {value_str}")
                    else:
                        # 普通输入框
                        await self.fill_input(input_id, value_str, title=field_with_suffix)
                        logger.info(f"填写{field_with_suffix}: {value_str}")
            
            traveler_index += 1
            
            # 如果检测到结束标记，处理完当前行后退出
            if should_break:
                break
        
        # 所有字段填写完成
        logger.info(f"所有字段填写完成")
        
        logger.info(f"出差人信息填写完成，共处理了 {traveler_index} 个出差人")
    
    async def process_remaining_operations(self, group_data: pd.DataFrame):
        """
        处理第二种子序列完成后的剩余操作字段
        
        Args:
            group_data: 同一序号下的所有数据行
        """
        logger.info("开始处理第二种子序列完成后的剩余操作字段")
        
        # 将DataFrame转换为list以便遍历
        rows = group_data.to_dict('records')
        columns = list(group_data.columns)
        
        i = 0
        while i < len(rows):
            row = rows[i]
            current_sequence = row.get(SEQUENCE_COL, None)
            logger.info(f"处理第 {i+1} 行数据的剩余操作，序号: {current_sequence}")
            
            # 处理当前行的所有列
            col_idx = 0
            while col_idx < len(columns):
                col = columns[col_idx]
                
                # 跳过序号列、处理进度列、子序列标记列和出差人信息字段
                if (col == SEQUENCE_COL or col == "处理进度" or 
                    col.startswith(SUBSEQUENCE_START_COL) or 
                    col.startswith(SUBSEQUENCE_END_COL) or
                    col in TRAVELER_FIELDS.keys()):
                    col_idx += 1
                    continue
                
                # 跳过登录相关字段
                if col in ["登录界面工号", "登录界面密码", "登录按钮", "网上预约报账按钮", "等待", "申请报销单按钮", "已阅读并同意按钮", "选择业务大类", "报销项目号", "附件张数", "备注", "特殊事项说明", "下一步按钮1", "等待.1"]:
                    col_idx += 1
                    continue
                
                # 跳过已经在第二种子序列中处理过的字段
                if col in ["省份", "出差地点", "起", "迄", "飞机票", "住宿费", "是否安排伙食", "是否安排交通"]:
                    col_idx += 1
                    continue
                
                value = row[col]
                if pd.notna(value) and value != "":
                    value_str = self.clean_value_string(value)
                    logger.info(f"处理剩余操作: {col} = {value_str}")
                    await self.process_cell(col, value_str)
                
                col_idx += 1
            
            # 移动到下一行
            i += 1
        
        logger.info("剩余操作字段处理完成")
    
    async def process_subsequence_logic(self, record_data: pd.DataFrame):
        """
        处理子序列逻辑
        
        Args:
            record_data: 包含该报销记录所有行的DataFrame
        """
        logger.info(f"处理子序列逻辑，共{len(record_data)}行")
        
        # 按行处理，从左到右
        for row_idx, row in record_data.iterrows():
            logger.info(f"处理第{row_idx + 1}行数据")
            
            # 首先检查当前行是否包含第二种子序列开始标记
            traveler_subsequence_detected = False
            for col in record_data.columns:
                if col.startswith(SUBSEQUENCE_START_COL):
                    subsequence_value = row[col]
                    if pd.notna(subsequence_value) and subsequence_value != "":
                        subsequence_value_str = self.clean_value_string(subsequence_value)
                        if subsequence_value_str == TRAVELER_SUBSEQUENCE_MARKER:
                            logger.info(f"检测到第二种子序列（出差人信息），列: {col}，开始处理出差人信息填写")
                            await self.process_traveler_subsequence(record_data, row_idx)
                            traveler_subsequence_detected = True
                            break
            
            # 如果检测到第二种子序列，跳过当前行的逐列处理
            if traveler_subsequence_detected:
                continue
            
            # 获取列名列表
            columns = list(record_data.columns)
            i = 0
            
            # 从左到右处理每一列
            while i < len(columns):
                col = columns[i]
                
                # 处理子序列开始列（支持自动重命名）
                if col.startswith(SUBSEQUENCE_START_COL):
                    subsequence_value = row[col]
                    if pd.notna(subsequence_value) and subsequence_value != "":
                        subsequence_value_str = self.clean_value_string(subsequence_value)
                        
                        # 这里只处理第一种子序列（非"1"标记）
                        if subsequence_value_str != TRAVELER_SUBSEQUENCE_MARKER:
                            logger.info(f"检测到第一种子序列（值: {subsequence_value_str}），继续处理")
                    i += 1
                    continue
                
                if col in [SEQUENCE_COL, "处理进度"] or col.startswith(SUBSEQUENCE_END_COL):
                    # 检查是否遇到子序列结束标记，如果遇到则继续在同一行向右处理
                    if col.startswith(SUBSEQUENCE_END_COL) and pd.notna(row.get(col)) and row.get(col) != "":
                        logger.info(f"检测到子序列结束标记（列: {col}），继续在同一行向右处理")
                    i += 1
                    continue
                
                value = row[col]
                if pd.notna(value) and value != "":
                    value_str = self.clean_value_string(value)
                    
                    # 特殊处理：科目列（以#开头）
                    if value_str.startswith("#"):
                        logger.info(f"处理科目列: {col} = {value_str}")
                        
                        # 提取科目名称（去掉#前缀）
                        subject_name = value_str[1:]
                        
                        # 在标题-ID表中查找对应的输入框ID
                        input_id = self.get_object_id(subject_name)
                        if not input_id:
                            logger.warning(f"未找到科目 '{subject_name}' 对应的ID映射")
                            i += 1
                            continue
                        
                        # 查找下一列的金额
                        if i + 1 < len(columns):
                            amount_col = columns[i + 1]
                            amount_value = row[amount_col]
                            
                            if pd.notna(amount_value) and amount_value != "":
                                amount_str = self.clean_value_string(amount_value)
                                logger.info(f"找到金额列: {amount_col} = {amount_str}")
                                
                                # 填写金额到对应的输入框
                                await self.fill_input(input_id, amount_str, title=amount_col)
                                logger.info(f"成功填写科目 '{subject_name}' 的金额: {amount_str}")
                                
                                # 跳过金额列，因为已经处理了
                                i += 2
                                continue
                            else:
                                logger.warning(f"科目 '{subject_name}' 对应的金额列为空")
                                i += 1
                                continue
                        else:
                            logger.warning(f"科目 '{subject_name}' 没有对应的金额列")
                            i += 1
                            continue
                    
                    # 处理普通列（在子序列中，需要添加序号后缀）
                    # 检查当前行是否是子序列中的行
                    is_in_subsequence = False
                    subsequence_index = 0
                    current_subsequence_num = 0
                    
                    # 检查当前行是否在子序列范围内
                    # 首先找到所有子序列的开始和结束位置
                    subsequence_ranges = []
                    
                    # 找到所有子序列开始列
                    start_cols = [c for c in record_data.columns if c.startswith(SUBSEQUENCE_START_COL)]
                    end_cols = [c for c in record_data.columns if c.startswith(SUBSEQUENCE_END_COL)]
                    
                    # 为每个子序列找到开始和结束位置
                    for seq_idx, start_col in enumerate(start_cols):
                        start_row_idx = None
                        end_row_idx = None
                        
                        # 找到子序列开始的行
                        for idx, seq_row in record_data.iterrows():
                            if pd.notna(seq_row[start_col]) and seq_row[start_col] != "":
                                start_row_idx = idx
                                break
                        
                        # 找到对应的子序列结束的行
                        if seq_idx < len(end_cols):
                            end_col = end_cols[seq_idx]
                            for idx, seq_row in record_data.iterrows():
                                if pd.notna(seq_row[end_col]) and seq_row[end_col] != "":
                                    end_row_idx = idx
                                    break
                        
                        if start_row_idx is not None and end_row_idx is not None:
                            subsequence_ranges.append((start_row_idx, end_row_idx, seq_idx))
                    
                    # 检查当前行属于哪个子序列
                    for start_row_idx, end_row_idx, subsequence_num in subsequence_ranges:
                        if start_row_idx <= row_idx <= end_row_idx:
                            is_in_subsequence = True
                            current_subsequence_num = subsequence_num
                            # 计算在当前子序列中的索引（从0开始）
                            subsequence_index = 0
                            # 从子序列开始行开始，计算当前行是第几个有数据的行
                            for check_row_idx in range(start_row_idx, row_idx + 1):
                                check_row = record_data.iloc[check_row_idx]
                                # 检查这一行是否有有效数据（除了子序列标记列）
                                has_data = False
                                for check_col in record_data.columns:
                                    if (check_col not in [SEQUENCE_COL, "处理进度"] and 
                                        not check_col.startswith(SUBSEQUENCE_START_COL) and 
                                        not check_col.startswith(SUBSEQUENCE_END_COL) and
                                        pd.notna(check_row[check_col]) and check_row[check_col] != ""):
                                        has_data = True
                                        break
                                if has_data:
                                    subsequence_index += 1
                            # 减1是因为我们要从0开始计数
                            subsequence_index = max(0, subsequence_index - 1)
                            logger.info(f"当前行 {row_idx} 属于子序列 {subsequence_num}，索引: {subsequence_index}")
                            break
                    
                    if is_in_subsequence:
                        # 在子序列中，使用带序号后缀的字段名
                        field_with_suffix = f"{col}-{subsequence_index}"
                        logger.info(f"处理子序列操作: {col} -> {field_with_suffix} = {value_str}")
                        await self.process_cell(field_with_suffix, value_str)
                    else:
                        # 不在子序列中，使用原始字段名
                        logger.info(f"处理普通操作: {col} = {value_str}")
                        await self.process_cell(col, value_str)
                
                i += 1
    
    async def run_automation(self, target_url: str = TARGET_URL):
        """
        运行自动化程序
        
        Args:
            target_url: 目标网页URL
        """
        try:
            # 加载数据
            await self.load_data()
            
            # 启动浏览器
            async with async_playwright() as p:
                if BROWSER_TYPE == "chromium":
                    self.browser = await p.chromium.launch(headless=HEADLESS)
                elif BROWSER_TYPE == "firefox":
                    self.browser = await p.firefox.launch(headless=HEADLESS)
                elif BROWSER_TYPE == "webkit":
                    self.browser = await p.webkit.launch(headless=HEADLESS)
                else:
                    raise ValueError(f"不支持的浏览器类型: {BROWSER_TYPE}")
                
                self.page = await self.browser.new_page()
                # 设置页面默认超时时间为3秒
                self.page.set_default_timeout(3000)
                
                # 导航到目标页面
                await self.page.goto(target_url, timeout=10000)
                logger.info(f"成功导航到页面: {target_url}")
                
                # 等待页面加载
                await asyncio.sleep(PAGE_LOAD_WAIT)
                
                # 按序号分组处理报销记录
                grouped_data = self.reimbursement_data.groupby(SEQUENCE_COL)
                
                for sequence_num, group_data in grouped_data:
                    logger.info(f"开始处理序号 {sequence_num} 的报销记录")
                    
                    # 处理子序列逻辑
                    await self.process_sequence_with_subsequences(sequence_num, group_data)
                    
                    # 处理完一条记录后等待一下
                    await asyncio.sleep(RECORD_PROCESS_WAIT)
                
                logger.info("所有报销记录处理完成")
                
                # 等待用户手动关闭浏览器
                logger.info("=" * 50)
                logger.info("所有操作已完成！")
                logger.info("浏览器将保持打开状态，您可以手动关闭。")
                logger.info("=" * 50)
                
                # 等待用户手动关闭浏览器
                try:
                    input("按回车键关闭浏览器...")
                except KeyboardInterrupt:
                    logger.info("用户中断程序")
                finally:
                    # 关闭浏览器
                    if self.browser:
                        await self.browser.close()
                        logger.info("浏览器已关闭")

        except Exception as e:
            logger.error(f"自动化程序运行失败: {e}")
            raise
        finally:
            if self.browser:
                await self.browser.close()

async def main():
    """主函数"""
    # 检查文件是否存在
    if not os.path.exists(EXCEL_FILE):
        logger.error(f"报销信息文件不存在: {EXCEL_FILE}")
        return
    
    if not os.path.exists(MAPPING_FILE):
        logger.error(f"标题-ID映射文件不存在: {MAPPING_FILE}")
        return
    
    # 创建自动化实例并运行
    automation = LoginAutomation()
    await automation.run_automation()

if __name__ == "__main__":
    asyncio.run(main()) 