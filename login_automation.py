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
            logger.warning(f"未找到标题 '{title}' 对应的ID映射")
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
        等待元素出现
        
        Args:
            element_id: 元素ID
            timeout: 超时时间（秒）
            
        Returns:
            是否成功找到元素
        """
        try:
            await self.page.wait_for_selector(f"#{element_id}", timeout=timeout * 1000)
            return True
        except TimeoutError:
            logger.warning(f"等待元素超时: {element_id}")
            return False
    
    async def fill_input(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        填写网页中的输入框
        
        Args:
            element_id: 输入框的ID
            value: 要填写的值
            retries: 重试次数
        """
        for attempt in range(retries):
            try:
                # 首先尝试在主页面查找
                if element_id and await self.wait_for_element(element_id):
                    await self.page.fill(f"#{element_id}", value)
                    logger.info(f"成功填写输入框 {element_id}: {value}")
                    return
                else:
                    # 如果主页面找不到，尝试在iframe中查找
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
                    
                    # 如果还是找不到，尝试通过name属性查找
                    try:
                        await self.page.fill(f"input[name='{element_id}']", value)
                        logger.info(f"通过name属性成功填写输入框 {element_id}: {value}")
                        return
                    except Exception as e:
                        logger.debug(f"通过name属性查找失败: {e}")
                        
                        # 在iframe中尝试通过name属性查找
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
                    
            except Exception as e:
                logger.warning(f"填写输入框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
        
        logger.error(f"填写输入框最终失败: {element_id}")
    
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
        
        # 方法1: 在所有iframe中动态查找按钮
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
            await asyncio.sleep(1)
            return True
        else:
            logger.error(f"点击按钮最终失败: {btnname}")
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
                # 首先尝试通过ID点击
                if element_id and await self.wait_for_element(element_id):
                    await self.page.click(f"#{element_id}")
                    logger.info(f"成功点击按钮: {element_id}")
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
                    # 缩短等待时间
                    await asyncio.sleep(2)
                    return True
                
                # 方法2: 通过JavaScript直接调用
                logger.info(f"尝试通过JavaScript调用: navToPrj('{value}')")
                try:
                    await self.page.evaluate(f"navToPrj('{value}')")
                    logger.info(f"成功点击导览框 (通过JavaScript): {value}")
                    # 缩短等待时间
                    await asyncio.sleep(2)
                    return True
                except Exception as js_error:
                    logger.debug(f"JavaScript调用失败: {js_error}")
                
                # 方法3: 通过文本内容查找
                logger.info(f"尝试通过文本内容查找: div:has-text('{value}')")
                text_selector = f"div:has-text('{value}')"
                
                if await self.page.locator(text_selector).count() > 0:
                    await self.page.click(text_selector)
                    logger.info(f"成功点击导览框 (通过文本): {value}")
                    # 缩短等待时间
                    await asyncio.sleep(2)
                    return True
                
                # 方法4: 通过title属性查找
                logger.info(f"尝试通过title属性查找: div[title*='{value}']")
                title_selector = f"div[title*='{value}']"
                
                if await self.page.locator(title_selector).count() > 0:
                    await self.page.click(title_selector)
                    logger.info(f"成功点击导览框 (通过title): {value}")
                    # 缩短等待时间
                    await asyncio.sleep(2)
                    return True
                
                # 方法5: 通过class和onclick组合查找
                logger.info(f"尝试通过class和onclick组合查找: div.syslink[onclick*='{value}']")
                class_selector = f"div.syslink[onclick*='{value}']"
                
                if await self.page.locator(class_selector).count() > 0:
                    await self.page.click(class_selector)
                    logger.info(f"成功点击导览框 (通过class+onclick): {value}")
                    # 缩短等待时间
                    await asyncio.sleep(2)
                    return True
                
                # 方法6: 通过第一个syslink元素查找
                logger.info("尝试通过第一个syslink元素查找")
                first_syslink = self.page.locator("div.syslink").first
                
                if await first_syslink.count() > 0:
                    await first_syslink.click()
                    logger.info(f"成功点击导览框 (通过第一个syslink): {value}")
                    # 缩短等待时间
                    await asyncio.sleep(2)
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
            
        element_id = self.get_object_id(title)
        if not element_id:
            return
            
        value_str = self.clean_value_string(value)
        
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
            await self.fill_input(element_id, value_str)
            
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
            
            # 等待银行卡选择弹窗出现
            logger.info("等待银行卡选择弹窗出现...")
            await asyncio.sleep(2)
            
            # 检查是否需要选择银行卡
            # 创建当前记录的DataFrame
            current_record = pd.DataFrame([{title: value_str}])
            await self.handle_bank_card_selection_for_transfer(value_str, current_record)
            return
        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith(BUTTON_PREFIX):
            button_value = value_str[len(BUTTON_PREFIX):]  # 去掉前缀符号
            await self.click_button(element_id)
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
        
        # 处理下拉框选择
        if title in DROPDOWN_FIELDS:
            # 获取下拉框的映射关系
            dropdown_mapping = DROPDOWN_FIELDS[title]
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
        
        # 处理普通输入框
        await self.fill_input(element_id, value_str)
    
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
                # 首先尝试在主页面查找
                if element_id and value and await self.wait_for_element(element_id):
                    # 选择对应的选项
                    await self.page.select_option(f"#{element_id}", value)
                    logger.info(f"成功选择下拉框 {element_id}: {value}")
                    await asyncio.sleep(ELEMENT_WAIT)
                    return
                else:
                    # 如果主页面找不到，尝试在iframe中查找
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
                    
                    # 如果还是找不到，尝试通过name属性查找
                    try:
                        await self.page.select_option(f"select[name='{element_id}']", value=value)
                        logger.info(f"通过name属性成功选择下拉框 {element_id}: {value}")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                    except Exception as e:
                        logger.debug(f"通过name属性查找失败: {e}")
                        
                        # 在iframe中尝试通过name属性查找
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
            
            # 等待更长时间，因为弹窗可能需要时间出现
            await asyncio.sleep(3)
            
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
            
            # 等待弹窗完全加载
            await asyncio.sleep(1)
            
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
            
            # 等待银行卡选择弹窗出现 - 使用更精确的选择器
            bank_dialog_found = False
            selectors_to_try = [
                "#paybankdiv",                           # 主要的银行卡选择弹窗ID
                "div.ui-dialog[aria-describedby='paybankdiv']",  # 完整的对话框
                "div.ui-dialog-title:has-text('请选择卡号')",  # 对话框标题
                "input[name='rdoacnt']",                 # 银行卡选择radio按钮
                "table[style*='background-color:#F2FAFD']",  # 银行卡表格
                "div.ui-dialog-content"                  # UI对话框内容
            ]
            
            # 方法1: 在主页面等待弹窗出现
            logger.info("方法1: 在主页面等待银行卡选择弹窗出现...")
            for selector in selectors_to_try:
                try:
                    await self.page.wait_for_selector(selector, timeout=5000)
                    logger.info(f"✓ 在主页面检测到银行卡选择弹窗，使用选择器: {selector}")
                    bank_dialog_found = True
                    break
                except Exception as e:
                    logger.debug(f"主页面选择器 {selector} 未找到: {e}")
                    continue
            
            # 方法2: 在所有iframe中查找弹窗
            if not bank_dialog_found:
                logger.info("方法2: 在所有iframe中查找银行卡选择弹窗...")
                frames = self.page.frames
                logger.info(f"找到 {len(frames)} 个iframe")
                
                for i, frame in enumerate(frames):
                    logger.info(f"检查iframe {i}: {frame.url}")
                    try:
                        for selector in selectors_to_try:
                            try:
                                await frame.wait_for_selector(selector, timeout=2000)
                                logger.info(f"✓ 在iframe {i} 中检测到银行卡选择弹窗，使用选择器: {selector}")
                                bank_dialog_found = True
                                # 将当前frame设置为活动frame
                                self.current_frame = frame
                                break
                            except Exception as e:
                                logger.debug(f"iframe {i} 选择器 {selector} 未找到: {e}")
                                continue
                        if bank_dialog_found:
                            break
                    except Exception as e:
                        logger.debug(f"检查iframe {i} 失败: {e}")
                        continue
            
            # 方法3: 如果方法1和2失败，尝试检测弹窗是否已经存在
            if not bank_dialog_found:
                logger.info("方法3: 尝试检测已存在的银行卡选择弹窗...")
                # 在主页面检测
                for selector in selectors_to_try:
                    try:
                        elements = await self.page.locator(selector).all()
                        if len(elements) > 0:
                            logger.info(f"✓ 在主页面检测到已存在的银行卡选择弹窗，使用选择器: {selector}")
                            bank_dialog_found = True
                            break
                    except Exception as e:
                        logger.debug(f"主页面检测选择器 {selector} 失败: {e}")
                        continue
                
                # 在iframe中检测
                if not bank_dialog_found:
                    frames = self.page.frames
                    for i, frame in enumerate(frames):
                        try:
                            for selector in selectors_to_try:
                                try:
                                    elements = await frame.locator(selector).all()
                                    if len(elements) > 0:
                                        logger.info(f"✓ 在iframe {i} 中检测到已存在的银行卡选择弹窗，使用选择器: {selector}")
                                        bank_dialog_found = True
                                        self.current_frame = frame
                                        break
                                except Exception as e:
                                    logger.debug(f"iframe {i} 检测选择器 {selector} 失败: {e}")
                                    continue
                            if bank_dialog_found:
                                break
                        except Exception as e:
                            logger.debug(f"检查iframe {i} 失败: {e}")
                            continue
            
            # 方法4: 检查是否有"请选择卡号"的标题
            if not bank_dialog_found:
                logger.info("方法4: 检查是否有'请选择卡号'标题...")
                # 在主页面检查
                try:
                    title_elements = await self.page.locator("text=请选择卡号").all()
                    if len(title_elements) > 0:
                        logger.info("✓ 在主页面检测到'请选择卡号'标题")
                        bank_dialog_found = True
                except Exception as e:
                    logger.debug(f"主页面检测'请选择卡号'标题失败: {e}")
                
                # 在iframe中检查
                if not bank_dialog_found:
                    frames = self.page.frames
                    for i, frame in enumerate(frames):
                        try:
                            title_elements = await frame.locator("text=请选择卡号").all()
                            if len(title_elements) > 0:
                                logger.info(f"✓ 在iframe {i} 中检测到'请选择卡号'标题")
                                bank_dialog_found = True
                                self.current_frame = frame
                                break
                        except Exception as e:
                            logger.debug(f"iframe {i} 检测'请选择卡号'标题失败: {e}")
                            continue
            
            # 方法5: 检查是否有radio按钮
            if not bank_dialog_found:
                logger.info("方法5: 检查是否有银行卡选择radio按钮...")
                # 在主页面检查
                try:
                    radio_elements = await self.page.locator("input[type='radio'][name='rdoacnt']").all()
                    if len(radio_elements) > 0:
                        logger.info(f"✓ 在主页面检测到 {len(radio_elements)} 个银行卡选择radio按钮")
                        bank_dialog_found = True
                except Exception as e:
                    logger.debug(f"主页面检测radio按钮失败: {e}")
                
                # 在iframe中检查
                if not bank_dialog_found:
                    frames = self.page.frames
                    for i, frame in enumerate(frames):
                        try:
                            radio_elements = await frame.locator("input[type='radio'][name='rdoacnt']").all()
                            if len(radio_elements) > 0:
                                logger.info(f"✓ 在iframe {i} 中检测到 {len(radio_elements)} 个银行卡选择radio按钮")
                                bank_dialog_found = True
                                self.current_frame = frame
                                break
                        except Exception as e:
                            logger.debug(f"iframe {i} 检测radio按钮失败: {e}")
                            continue
            
            # 方法6: 检查是否有ui-dialog类的元素
            if not bank_dialog_found:
                logger.info("方法6: 检查是否有ui-dialog类的元素...")
                # 在主页面检查
                try:
                    dialog_elements = await self.page.locator("div.ui-dialog").all()
                    if len(dialog_elements) > 0:
                        logger.info(f"✓ 在主页面检测到 {len(dialog_elements)} 个ui-dialog元素")
                        bank_dialog_found = True
                except Exception as e:
                    logger.debug(f"主页面检测ui-dialog元素失败: {e}")
                
                # 在iframe中检查
                if not bank_dialog_found:
                    frames = self.page.frames
                    for i, frame in enumerate(frames):
                        try:
                            dialog_elements = await frame.locator("div.ui-dialog").all()
                            if len(dialog_elements) > 0:
                                logger.info(f"✓ 在iframe {i} 中检测到 {len(dialog_elements)} 个ui-dialog元素")
                                bank_dialog_found = True
                                self.current_frame = frame
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
                    # 确定在哪个frame中操作
                    target_frame = getattr(self, 'current_frame', self.page)
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
            
            # 等待弹窗完全加载
            await asyncio.sleep(1)
            
            # 确定在哪个frame中操作
            target_frame = getattr(self, 'current_frame', self.page)
            
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
            
            # 首先在活动frame中查找（如果有的话）
            target_frame = getattr(self, 'current_frame', self.page)
            for selector in confirm_selectors:
                try:
                    confirm_button = target_frame.locator(selector).first
                    if await confirm_button.count() > 0:
                        await confirm_button.click()
                        logger.info(f"✓ 在活动frame中成功点击确定按钮 (使用选择器: {selector})")
                        await asyncio.sleep(1)
                        confirm_clicked = True
                        break
                except Exception as e:
                    logger.debug(f"活动frame确定按钮选择器 {selector} 失败: {e}")
                    continue
            
            # 如果活动frame中没找到，尝试在主页面查找
            if not confirm_clicked:
                for selector in confirm_selectors:
                    try:
                        confirm_button = self.page.locator(selector).first
                        if await confirm_button.count() > 0:
                            await confirm_button.click()
                            logger.info(f"✓ 在主页面成功点击确定按钮 (使用选择器: {selector})")
                            await asyncio.sleep(1)
                            confirm_clicked = True
                            break
                    except Exception as e:
                        logger.debug(f"主页面确定按钮选择器 {selector} 失败: {e}")
                        continue
            
            # 如果主页面也没找到，尝试在所有iframe中查找
            if not confirm_clicked:
                frames = self.page.frames
                for i, frame in enumerate(frames):
                    for selector in confirm_selectors:
                        try:
                            confirm_button = frame.locator(selector).first
                            if await confirm_button.count() > 0:
                                await confirm_button.click()
                                logger.info(f"✓ 在iframe {i} 中成功点击确定按钮 (使用选择器: {selector})")
                                await asyncio.sleep(1)
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
            await self.handle_login_with_captcha(group_data)
        else:
            # 处理子序列逻辑
            await self.process_subsequences(group_data)
    
    async def process_subsequences(self, group_data: pd.DataFrame):
        """
        处理子序列逻辑：从子序列开始到子序列结束，然后跳转到下一行相同序号的子序列开始
        
        Args:
            group_data: 同一序号下的所有数据行
        """
        # 将DataFrame转换为list以便遍历
        rows = group_data.to_dict('records')
        columns = list(group_data.columns)
        
        i = 0
        while i < len(rows):
            row = rows[i]
            logger.info(f"处理第 {i+1} 行数据")
            
            # 查找子序列开始列的位置
            subsequence_start_idx = None
            subsequence_end_idx = None
            
            try:
                subsequence_start_idx = columns.index(SUBSEQUENCE_START_COL)
            except ValueError:
                # 如果没有子序列开始列，按普通方式处理这一行
                await self.process_single_row(row, columns)
                i += 1
                continue
            
            try:
                subsequence_end_idx = columns.index(SUBSEQUENCE_END_COL)
            except ValueError:
                # 如果没有子序列结束列，按普通方式处理这一行
                await self.process_single_row(row, columns)
                i += 1
                continue
            
            # 处理从子序列开始到子序列结束的列
            col_idx = 0
            while col_idx < len(columns):
                col = columns[col_idx]
                
                # 跳过序号列
                if col == SEQUENCE_COL:
                    col_idx += 1
                    continue
                
                # 如果到达子序列开始列，开始处理子序列
                if col_idx == subsequence_start_idx:
                    # 处理从子序列开始到子序列结束的所有列
                    for subseq_col_idx in range(subsequence_start_idx, min(subsequence_end_idx + 1, len(columns))):
                        subseq_col = columns[subseq_col_idx]
                        if subseq_col not in [SUBSEQUENCE_START_COL, SUBSEQUENCE_END_COL]:
                            value = row[subseq_col]
                            if pd.notna(value) and value != "":
                                value_str = self.clean_value_string(value)
                                logger.info(f"处理子序列操作: {subseq_col} = {value_str}")
                                await self.process_cell(subseq_col, value_str)
                    
                    # 子序列处理完成，跳转到子序列结束后
                    col_idx = subsequence_end_idx + 1
                    break
                else:
                    # 处理子序列开始之前的列
                    value = row[col]
                    if pd.notna(value) and value != "":
                        value_str = self.clean_value_string(value)
                        logger.info(f"处理普通操作: {col} = {value_str}")
                        await self.process_cell(col, value_str)
                    col_idx += 1
            
            i += 1
    
    async def process_single_row(self, row: Dict, columns: List[str]):
        """
        处理单行数据（没有子序列逻辑）
        
        Args:
            row: 行数据字典
            columns: 列名列表
        """
        for col in columns:
            if col == SEQUENCE_COL:  # 跳过序号列
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
            
            # 跳过登录相关的列和序号列
            if col in ["序号", "登录界面工号", "登录界面密码", "登录按钮"]:
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
                            await self.fill_input(input_id, amount_str)
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
        
        logger.info(f"开始处理序号 {sequence_num} 的报销记录，共{len(record_data)}行数据")
        
        # 检查是否有子序列列
        has_subsequence = (SUBSEQUENCE_START_COL in record_data.columns and 
                          SUBSEQUENCE_END_COL in record_data.columns)
        
        if has_subsequence:
            # 处理子序列逻辑
            await self.process_subsequence_logic(record_data)
        else:
            # 处理普通逻辑（假设只有一行数据）
            row = record_data.iloc[0]
            for col in record_data.columns:
                if col == SEQUENCE_COL:
                    continue
                
                value = row[col]
                if pd.notna(value) and value != "":
                    await self.process_cell(col, value)
        
        logger.info(f"序号 {sequence_num} 的报销记录处理完成")
    
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
            
            # 从左到右处理每一列
            for col in record_data.columns:
                if col in [SEQUENCE_COL, SUBSEQUENCE_START_COL, SUBSEQUENCE_END_COL]:
                    continue
                
                value = row[col]
                if pd.notna(value) and value != "":
                    await self.process_cell(col, value)
            
            # 如果遇到子序列结束，等待一下再继续
            if pd.notna(row.get(SUBSEQUENCE_END_COL, pd.NA)):
                logger.info("检测到子序列结束，等待处理下一行")
                await asyncio.sleep(0.3)
    
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
                
                # 导航到目标页面
                await self.page.goto(target_url)
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
    asyncio.run(main()) 