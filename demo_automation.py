import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import logging
import os

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DemoAutomation:
    def __init__(self):
        self.title_id_mapping = {}
        self.reimbursement_data = None
        
    def load_data(self):
        """加载Excel数据和标题-ID映射"""
        try:
            # 加载标题-ID映射
            mapping_df = pd.read_excel("标题-ID.xlsx")
            self.title_id_mapping = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))
            logger.info(f"成功加载标题-ID映射，共{len(self.title_id_mapping)}条记录")
            
            # 加载报销信息数据
            self.reimbursement_data = pd.read_excel("报销信息.xlsx", sheet_name="BaoXiao_sheet")
            logger.info(f"成功加载报销信息数据，共{len(self.reimbursement_data)}行")
            
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
    
    async def fill_input(self, page, element_id: str, value: str):
        """在输入框中填写内容"""
        try:
            if element_id and value:
                await page.fill(f"#{element_id}", str(value))
                logger.info(f"成功填写输入框 {element_id}: {value}")
                await asyncio.sleep(0.5)
        except Exception as e:
            logger.error(f"填写输入框失败 {element_id}: {e}")
    
    async def select_dropdown(self, page, element_id: str, value: str):
        """选择下拉框中的选项"""
        try:
            if element_id and value:
                await page.select_option(f"#{element_id}", value)
                logger.info(f"成功选择下拉框 {element_id}: {value}")
                await asyncio.sleep(0.5)
        except Exception as e:
            logger.error(f"选择下拉框失败 {element_id}: {e}")
    
    async def click_button(self, page, element_id: str):
        """点击网页中的按钮"""
        try:
            if element_id:
                await page.click(f"#{element_id}")
                logger.info(f"成功点击按钮: {element_id}")
                await asyncio.sleep(0.5)
        except Exception as e:
            logger.error(f"点击按钮失败 {element_id}: {e}")
    
    async def process_cell(self, page, title: str, value):
        """处理单个单元格的内容"""
        if pd.isna(value) or value == "":
            return
            
        element_id = self.get_object_id(title)
        if not element_id:
            return
            
        value_str = str(value).strip()
        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith('$'):
            await self.click_button(page, element_id)
            return
        
        # 处理下拉框选择
        dropdown_fields = ["支付方式", "报销类型", "部门"]
        if title in dropdown_fields:
            await self.select_dropdown(page, element_id, value_str)
            return
        
        # 处理普通输入框
        await self.fill_input(page, element_id, value_str)
    
    async def process_record(self, page, record_data: pd.DataFrame):
        """处理单条报销记录"""
        sequence_num = record_data["序号"].iloc[0]
        logger.info(f"开始处理序号 {sequence_num} 的报销记录")
        
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
                        await self.process_cell(page, col, value)
                
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
                    await self.process_cell(page, col, value)
        
        logger.info(f"序号 {sequence_num} 的报销记录处理完成")
    
    async def run_demo(self):
        """运行演示程序"""
        try:
            # 加载数据
            self.load_data()
            
            # 启动浏览器
            logger.info("启动浏览器...")
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False)
                logger.info("浏览器启动成功")
                
                page = await browser.new_page()
                logger.info("新页面创建成功")
                
                # 导航到测试页面
                test_url = "file:///C:/Users/FH/PycharmProjects/CursorCode8-5/test_page.html"
                await page.goto(test_url)
                logger.info(f"成功导航到页面: {test_url}")
                
                # 等待页面加载
                await asyncio.sleep(2)
                
                # 按序号分组处理报销记录
                grouped_data = self.reimbursement_data.groupby("序号")
                
                for sequence_num, group_data in grouped_data:
                    logger.info(f"开始处理序号 {sequence_num} 的报销记录")
                    await self.process_record(page, group_data)
                    
                    # 处理完一条记录后等待一下
                    await asyncio.sleep(2)
                    
                    # 尝试提交表单
                    try:
                        await page.click("#submit_btn")
                        logger.info("成功点击提交按钮")
                        await asyncio.sleep(2)
                    except:
                        logger.warning("未找到提交按钮或提交失败")
                
                logger.info("所有报销记录处理完成")
                
                # 等待用户查看结果
                logger.info("演示完成，等待5秒后关闭浏览器...")
                await asyncio.sleep(5)
                
                await browser.close()
                logger.info("浏览器已关闭")
                
        except Exception as e:
            logger.error(f"演示程序运行失败: {e}")
            import traceback
            traceback.print_exc()

async def main():
    """主函数"""
    # 检查文件是否存在
    if not os.path.exists("报销信息.xlsx"):
        logger.error("报销信息文件不存在")
        return
    
    if not os.path.exists("标题-ID.xlsx"):
        logger.error("标题-ID映射文件不存在")
        return
    
    # 创建演示实例并运行
    demo = DemoAutomation()
    await demo.run_demo()

if __name__ == "__main__":
    asyncio.run(main()) 