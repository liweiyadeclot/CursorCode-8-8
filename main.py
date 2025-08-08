import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import logging
from typing import Optional, Dict, Any
import os

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ReimbursementAutomation:
    def __init__(self, excel_file: str, mapping_file: str, sheet_name: str = "BaoXiao_sheet"):
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
    
    async def click_button(self, element_id: str):
        """
        点击网页中的按钮
        
        Args:
            element_id: 按钮的ID
        """
        try:
            if element_id:
                await self.page.click(f"#{element_id}")
                logger.info(f"成功点击按钮: {element_id}")
                await asyncio.sleep(0.5)  # 等待页面响应
        except Exception as e:
            logger.error(f"点击按钮失败 {element_id}: {e}")
    
    async def fill_input(self, element_id: str, value: str):
        """
        在输入框中填写内容
        
        Args:
            element_id: 输入框的ID
            value: 要填写的内容
        """
        try:
            if element_id and value:
                await self.page.fill(f"#{element_id}", str(value))
                logger.info(f"成功填写输入框 {element_id}: {value}")
                await asyncio.sleep(0.3)
        except Exception as e:
            logger.error(f"填写输入框失败 {element_id}: {e}")
    
    async def select_dropdown(self, element_id: str, value: str):
        """
        选择下拉框中的选项
        
        Args:
            element_id: 下拉框的ID
            value: 要选择的选项值
        """
        try:
            if element_id and value:
                # 先点击下拉框
                await self.page.click(f"#{element_id}")
                await asyncio.sleep(0.3)
                
                # 选择对应的选项
                await self.page.select_option(f"#{element_id}", value)
                logger.info(f"成功选择下拉框 {element_id}: {value}")
                await asyncio.sleep(0.3)
        except Exception as e:
            logger.error(f"选择下拉框失败 {element_id}: {e}")
    
    async def process_cell(self, title: str, value: Any):
        """
        处理单个单元格的内容
        
        Args:
            title: 列标题
            value: 单元格值
        """
        if pd.isna(value):
            return
            
        element_id = self.get_object_id(title)
        if not element_id:
            return
            
        value_str = str(value).strip()
        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith('$'):
            button_value = value_str[1:]  # 去掉$符号
            await self.click_button(element_id)
            return
        
        # 处理下拉框选择
        if title in ["支付方式", "报销类型", "部门"]:  # 根据实际需要调整
            await self.select_dropdown(element_id, value_str)
            return
        
        # 处理普通输入框
        await self.fill_input(element_id, value_str)
    
    async def process_reimbursement_record(self, record_data: pd.DataFrame):
        """
        处理单条报销记录
        
        Args:
            record_data: 包含该报销记录所有行的DataFrame
        """
        logger.info(f"开始处理报销记录，共{len(record_data)}行数据")
        
        # 检查是否有子序列列
        has_subsequence = "子序列开始" in record_data.columns and "子序列结束" in record_data.columns
        
        if has_subsequence:
            # 处理子序列逻辑
            for _, row in record_data.iterrows():
                # 从左到右处理每一行
                for col in record_data.columns:
                    if col in ["序号", "子序列开始", "子序列结束"]:
                        continue
                    
                    value = row[col]
                    if pd.notna(value):
                        await self.process_cell(col, value)
                
                # 如果遇到子序列结束，等待一下再继续
                if pd.notna(row.get("子序列结束", pd.NA)):
                    await asyncio.sleep(1)
        else:
            # 处理普通逻辑（假设只有一行数据）
            row = record_data.iloc[0]
            for col in record_data.columns:
                if col == "序号":
                    continue
                
                value = row[col]
                if pd.notna(value):
                    await self.process_cell(col, value)
    
    async def run_automation(self, target_url: str):
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
                self.browser = await p.chromium.launch(headless=False)  # 设置为True可以隐藏浏览器
                self.page = await self.browser.new_page()
                
                # 导航到目标页面
                await self.page.goto(target_url)
                logger.info(f"成功导航到页面: {target_url}")
                
                # 等待页面加载
                await asyncio.sleep(2)
                
                # 按序号分组处理报销记录
                if "序号" in self.reimbursement_data.columns:
                    grouped_data = self.reimbursement_data.groupby("序号")
                    
                    for sequence_num, group_data in grouped_data:
                        logger.info(f"开始处理序号 {sequence_num} 的报销记录")
                        await self.process_reimbursement_record(group_data)
                        
                        # 处理完一条记录后等待一下
                        await asyncio.sleep(2)
                        
                        # 这里可以添加提交表单的逻辑
                        # await self.submit_form()
                        
                else:
                    logger.warning("未找到序号列，将处理所有数据作为单条记录")
                    await self.process_reimbursement_record(self.reimbursement_data)
                
                logger.info("所有报销记录处理完成")
                
        except Exception as e:
            logger.error(f"自动化程序运行失败: {e}")
            raise
        finally:
            if self.browser:
                await self.browser.close()

async def main():
    """主函数"""
    # 配置文件路径
    excel_file = "报销信息.xlsx"
    mapping_file = "标题-ID.xlsx"
    target_url = "http://your-reimbursement-system-url.com"  # 替换为实际的URL
    
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        logger.error(f"报销信息文件不存在: {excel_file}")
        return
    
    if not os.path.exists(mapping_file):
        logger.error(f"标题-ID映射文件不存在: {mapping_file}")
        return
    
    # 创建自动化实例并运行
    automation = ReimbursementAutomation(excel_file, mapping_file)
    await automation.run_automation(target_url)

if __name__ == "__main__":
    asyncio.run(main())
