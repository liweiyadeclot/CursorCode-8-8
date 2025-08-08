import asyncio
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError
import logging
from typing import Optional, Dict, Any, List
import os
import time
from config import *

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

class ReimbursementAutomation:
    def __init__(self, excel_file: str = EXCEL_FILE, mapping_file: str = MAPPING_FILE, 
                 sheet_name: str = SHEET_NAME):
        """
        初始化报销自动化类
        
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
    
    async def wait_for_element(self, element_id: str, timeout: int = 10) -> bool:
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
    
    async def click_button(self, element_id: str, retries: int = MAX_RETRIES):
        """
        点击网页中的按钮
        
        Args:
            element_id: 按钮的ID
            retries: 重试次数
        """
        for attempt in range(retries):
            try:
                if element_id and await self.wait_for_element(element_id):
                    await self.page.click(f"#{element_id}")
                    logger.info(f"成功点击按钮: {element_id}")
                    await asyncio.sleep(BUTTON_CLICK_WAIT)
                    return
                else:
                    logger.warning(f"按钮元素不存在: {element_id}")
                    return
            except Exception as e:
                logger.warning(f"点击按钮失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"点击按钮最终失败: {element_id}")
    
    async def fill_input(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        在输入框中填写内容
        
        Args:
            element_id: 输入框的ID
            value: 要填写的内容
            retries: 重试次数
        """
        for attempt in range(retries):
            try:
                if element_id and value and await self.wait_for_element(element_id):
                    # 先清空输入框
                    await self.page.fill(f"#{element_id}", "")
                    await asyncio.sleep(0.1)
                    
                    # 填写新内容
                    await self.page.fill(f"#{element_id}", str(value))
                    logger.info(f"成功填写输入框 {element_id}: {value}")
                    await asyncio.sleep(ELEMENT_WAIT)
                    return
                else:
                    logger.warning(f"输入框元素不存在: {element_id}")
                    return
            except Exception as e:
                logger.warning(f"填写输入框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"填写输入框最终失败: {element_id}")
    
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
                if element_id and value and await self.wait_for_element(element_id):
                    # 选择对应的选项
                    await self.page.select_option(f"#{element_id}", value)
                    logger.info(f"成功选择下拉框 {element_id}: {value}")
                    await asyncio.sleep(ELEMENT_WAIT)
                    return
                else:
                    logger.warning(f"下拉框元素不存在: {element_id}")
                    return
            except Exception as e:
                logger.warning(f"选择下拉框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"选择下拉框最终失败: {element_id}")
    
    async def click_navigation_panel(self, element_id: str, value: str, retries: int = MAX_RETRIES):
        """
        点击系统导览框
        
        Args:
            element_id: 导览框的ID或选择器
            value: 要点击的导览框标识（如WF_YB6）
            retries: 重试次数
        """
        for attempt in range(retries):
            try:
                if element_id and value:
                    # 方法1: 通过ID直接点击
                    try:
                        if await self.wait_for_element(element_id):
                            await self.page.click(f"#{element_id}")
                            logger.info(f"成功点击导览框 {element_id}")
                            await asyncio.sleep(ELEMENT_WAIT)
                            return
                    except:
                        pass
                    
                    # 方法2: 通过onclick属性查找并点击
                    try:
                        selector = f"div[onclick*='{value}']"
                        await self.page.wait_for_selector(selector, timeout=2000)
                        await self.page.click(selector)
                        logger.info(f"成功点击导览框 (通过onclick): {value}")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                    except:
                        pass
                    
                    # 方法3: 通过title属性查找并点击
                    try:
                        selector = f"div[title*='{value}']"
                        await self.page.wait_for_selector(selector, timeout=2000)
                        await self.page.click(selector)
                        logger.info(f"成功点击导览框 (通过title): {value}")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                    except:
                        pass
                    
                    # 方法4: 通过class和文本内容查找
                    try:
                        selector = f"div.syslink:has-text('{value}')"
                        await self.page.wait_for_selector(selector, timeout=2000)
                        await self.page.click(selector)
                        logger.info(f"成功点击导览框 (通过class和文本): {value}")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                    except:
                        pass
                    
                    # 方法5: 通过JavaScript执行onclick
                    try:
                        await self.page.evaluate(f"navToPrj('{value}')")
                        logger.info(f"成功执行导览框JavaScript: {value}")
                        await asyncio.sleep(ELEMENT_WAIT)
                        return
                    except:
                        pass
                    
                    logger.warning(f"无法找到或点击导览框: {element_id}, 值: {value}")
                    return
                    
            except Exception as e:
                logger.warning(f"点击导览框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(RETRY_DELAY)
                else:
                    logger.error(f"点击导览框最终失败: {element_id}")

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
            
        value_str = str(value).strip()
        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith(BUTTON_PREFIX):
            button_value = value_str[len(BUTTON_PREFIX):]  # 去掉前缀符号
            await self.click_button(element_id)
            return
        
        # 处理系统导览框点击操作（以@开头）
        if value_str.startswith('@'):
            nav_value = value_str[1:]  # 去掉@符号
            await self.click_navigation_panel(element_id, nav_value)
            return
        
        # 处理下拉框选择
        if title in DROPDOWN_FIELDS:
            await self.select_dropdown(element_id, value_str)
            return
        
        # 处理普通输入框
        await self.fill_input(element_id, value_str)
    
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
                await asyncio.sleep(1)
    
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
    
    async def submit_form(self):
        """提交表单"""
        try:
            # 查找提交按钮（需要根据实际页面调整）
            submit_selectors = [
                "button[type='submit']",
                "input[type='submit']",
                "#submit",
                "#save",
                ".submit-btn"
            ]
            
            for selector in submit_selectors:
                try:
                    if await self.page.wait_for_selector(selector, timeout=2000):
                        await self.page.click(selector)
                        logger.info(f"成功点击提交按钮: {selector}")
                        await asyncio.sleep(2)
                        return True
                except:
                    continue
            
            logger.warning("未找到提交按钮")
            return False
            
        except Exception as e:
            logger.error(f"提交表单失败: {e}")
            return False
    
    async def validate_form(self) -> bool:
        """验证表单填写是否正确"""
        try:
            # 这里可以添加表单验证逻辑
            # 例如检查必填字段是否已填写等
            logger.info("表单验证通过")
            return True
        except Exception as e:
            logger.error(f"表单验证失败: {e}")
            return False
    
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
                    await self.process_reimbursement_record(group_data)
                    
                    # 处理完一条记录后等待一下
                    await asyncio.sleep(RECORD_PROCESS_WAIT)
                    
                    # 验证表单
                    if VALIDATE_BEFORE_SUBMIT:
                        if await self.validate_form():
                            # 提交表单
                            await self.submit_form()
                        else:
                            logger.error(f"序号 {sequence_num} 的表单验证失败，跳过提交")
                
                logger.info("所有报销记录处理完成")
                
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
    automation = ReimbursementAutomation()
    await automation.run_automation()

if __name__ == "__main__":
    asyncio.run(main()) 