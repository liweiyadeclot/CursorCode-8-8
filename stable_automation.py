import asyncio
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError
import logging
from typing import Optional, Dict, Any, List
import os
import time

# 配置日志 - 简化版本
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ReimbursementAutomation:
    def __init__(self, excel_file: str = "报销信息.xlsx", mapping_file: str = "标题-ID.xlsx", 
                 sheet_name: str = "BaoXiao_sheet"):
        """
        初始化报销自动化类
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
            if "序号" not in self.reimbursement_data.columns:
                raise ValueError("缺少必要的列: 序号")
                
        except Exception as e:
            logger.error(f"加载数据失败: {e}")
            raise
    
    def get_object_id(self, title: str) -> str:
        """根据表头标题获取对应的网页object id"""
        if title in self.title_id_mapping:
            return self.title_id_mapping[title]
        else:
            logger.warning(f"未找到标题 '{title}' 对应的ID映射")
            return ""
    
    async def wait_for_element(self, element_id: str, timeout: int = 10) -> bool:
        """等待元素出现"""
        try:
            await self.page.wait_for_selector(f"#{element_id}", timeout=timeout * 1000)
            return True
        except TimeoutError:
            logger.warning(f"等待元素超时: {element_id}")
            return False
    
    async def click_button(self, element_id: str, retries: int = 3):
        """点击网页中的按钮"""
        for attempt in range(retries):
            try:
                if element_id and await self.wait_for_element(element_id):
                    await self.page.click(f"#{element_id}")
                    logger.info(f"成功点击按钮: {element_id}")
                    await asyncio.sleep(0.5)
                    return
                else:
                    logger.warning(f"按钮元素不存在: {element_id}")
                    return
            except Exception as e:
                logger.warning(f"点击按钮失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(1)
                else:
                    logger.error(f"点击按钮最终失败: {element_id}")
    
    async def fill_input(self, element_id: str, value: str, retries: int = 3):
        """在输入框中填写内容"""
        for attempt in range(retries):
            try:
                if element_id and value and await self.wait_for_element(element_id):
                    # 先清空输入框
                    await self.page.fill(f"#{element_id}", "")
                    await asyncio.sleep(0.1)
                    
                    # 填写新内容
                    await self.page.fill(f"#{element_id}", str(value))
                    logger.info(f"成功填写输入框 {element_id}: {value}")
                    await asyncio.sleep(0.3)
                    return
                else:
                    logger.warning(f"输入框元素不存在: {element_id}")
                    return
            except Exception as e:
                logger.warning(f"填写输入框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(1)
                else:
                    logger.error(f"填写输入框最终失败: {element_id}")
    
    async def select_dropdown(self, element_id: str, value: str, retries: int = 3):
        """选择下拉框中的选项"""
        for attempt in range(retries):
            try:
                if element_id and value and await self.wait_for_element(element_id):
                    # 选择对应的选项
                    await self.page.select_option(f"#{element_id}", value)
                    logger.info(f"成功选择下拉框 {element_id}: {value}")
                    await asyncio.sleep(0.3)
                    return
                else:
                    logger.warning(f"下拉框元素不存在: {element_id}")
                    return
            except Exception as e:
                logger.warning(f"选择下拉框失败 (尝试 {attempt + 1}/{retries}): {element_id} - {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(1)
                else:
                    logger.error(f"选择下拉框最终失败: {element_id}")
    
    async def process_cell(self, title: str, value: Any):
        """处理单个单元格的内容"""
        if pd.isna(value) or value == "":
            return
            
        element_id = self.get_object_id(title)
        if not element_id:
            return
            
        value_str = str(value).strip()
        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith('$'):
            button_value = value_str[1:]  # 去掉前缀符号
            await self.click_button(element_id)
            return
        
        # 处理下拉框选择
        dropdown_fields = ["支付方式", "报销类型", "部门"]
        if title in dropdown_fields:
            await self.select_dropdown(element_id, value_str)
            return
        
        # 处理普通输入框
        await self.fill_input(element_id, value_str)
    
    async def process_reimbursement_record(self, record_data: pd.DataFrame):
        """处理单条报销记录"""
        sequence_num = record_data["序号"].iloc[0]
        self.current_sequence = sequence_num
        
        logger.info(f"开始处理序号 {sequence_num} 的报销记录，共{len(record_data)}行数据")
        
        # 检查是否有子序列列
        has_subsequence = ("子序列开始" in record_data.columns and 
                          "子序列结束" in record_data.columns)
        
        if has_subsequence:
            # 处理子序列逻辑
            logger.info(f"处理子序列逻辑，共{len(record_data)}行")
            
            # 按行处理，从左到右
            for row_idx, row in record_data.iterrows():
                logger.info(f"处理第{row_idx + 1}行数据")
                
                # 从左到右处理每一列
                for col in record_data.columns:
                    if col in ["序号", "子序列开始", "子序列结束"]:
                        continue
                    
                    value = row[col]
                    if pd.notna(value) and value != "":
                        await self.process_cell(col, value)
                
                # 如果遇到子序列结束，等待一下再继续
                if pd.notna(row.get("子序列结束", pd.NA)):
                    logger.info("检测到子序列结束，等待处理下一行")
                    await asyncio.sleep(1)
        else:
            # 处理普通逻辑（假设只有一行数据）
            row = record_data.iloc[0]
            for col in record_data.columns:
                if col == "序号":
                    continue
                
                value = row[col]
                if pd.notna(value) and value != "":
                    await self.process_cell(col, value)
        
        logger.info(f"序号 {sequence_num} 的报销记录处理完成")
    
    async def submit_form(self):
        """提交表单"""
        try:
            # 查找提交按钮
            submit_selectors = [
                "button[type='submit']",
                "input[type='submit']",
                "#submit_btn",
                "#save_btn",
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
    
    async def run_automation(self, target_url: str = "file:///C:/Users/FH/PycharmProjects/CursorCode8-5/test_page.html"):
        """运行自动化程序"""
        try:
            # 加载数据
            await self.load_data()
            
            # 启动浏览器
            logger.info("启动浏览器...")
            async with async_playwright() as p:
                self.browser = await p.chromium.launch(headless=False)
                logger.info("浏览器启动成功")
                
                self.page = await self.browser.new_page()
                logger.info("新页面创建成功")
                
                # 导航到目标页面
                await self.page.goto(target_url)
                logger.info(f"成功导航到页面: {target_url}")
                
                # 等待页面加载
                await asyncio.sleep(2)
                
                # 按序号分组处理报销记录
                grouped_data = self.reimbursement_data.groupby("序号")
                
                for sequence_num, group_data in grouped_data:
                    logger.info(f"开始处理序号 {sequence_num} 的报销记录")
                    await self.process_reimbursement_record(group_data)
                    
                    # 处理完一条记录后等待一下
                    await asyncio.sleep(2)
                    
                    # 提交表单
                    await self.submit_form()
                
                logger.info("所有报销记录处理完成")
                
        except Exception as e:
            logger.error(f"自动化程序运行失败: {e}")
            raise
        finally:
            if self.browser:
                await self.browser.close()
                logger.info("浏览器已关闭")

async def main():
    """主函数"""
    # 检查文件是否存在
    if not os.path.exists("报销信息.xlsx"):
        logger.error("报销信息文件不存在")
        return
    
    if not os.path.exists("标题-ID.xlsx"):
        logger.error("标题-ID映射文件不存在")
        return
    
    # 创建自动化实例并运行
    automation = ReimbursementAutomation()
    await automation.run_automation()

if __name__ == "__main__":
    asyncio.run(main()) 