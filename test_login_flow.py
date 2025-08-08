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

class TestLoginFlow:
    def __init__(self):
        self.title_id_mapping = {}
        self.reimbursement_data = None
        self.browser = None
        self.page = None
        
    async def load_data(self):
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
    
    async def wait_for_element(self, element_id: str, timeout: int = 10) -> bool:
        """等待元素出现"""
        try:
            await self.page.wait_for_selector(f"#{element_id}", timeout=timeout * 1000)
            return True
        except TimeoutError:
            logger.warning(f"等待元素超时: {element_id}")
            return False
    
    async def fill_input(self, element_id: str, value: str):
        """在输入框中填写内容"""
        try:
            if element_id and value and await self.wait_for_element(element_id):
                await self.page.fill(f"#{element_id}", str(value))
                logger.info(f"成功填写输入框 {element_id}: {value}")
                await asyncio.sleep(0.5)
        except Exception as e:
            logger.error(f"填写输入框失败 {element_id}: {e}")
    
    async def click_button(self, element_id: str):
        """点击网页中的按钮"""
        try:
            if element_id and await self.wait_for_element(element_id):
                await self.page.click(f"#{element_id}")
                logger.info(f"成功点击按钮: {element_id}")
                await asyncio.sleep(1)
        except Exception as e:
            logger.error(f"点击按钮失败 {element_id}: {e}")
    
    async def click_navigation_panel(self, element_id: str, value: str):
        """点击系统导览框"""
        try:
            if element_id and value:
                # 尝试通过JavaScript执行onclick
                try:
                    await self.page.evaluate(f"navToPrj('{value}')")
                    logger.info(f"成功执行导览框JavaScript: {value}")
                    await asyncio.sleep(0.5)
                    return
                except:
                    pass
                
                # 尝试通过onclick属性查找并点击
                try:
                    selector = f"div[onclick*='{value}']"
                    await self.page.wait_for_selector(selector, timeout=2000)
                    await self.page.click(selector)
                    logger.info(f"成功点击导览框 (通过onclick): {value}")
                    await asyncio.sleep(0.5)
                    return
                except:
                    pass
                
                logger.warning(f"无法找到或点击导览框: {element_id}, 值: {value}")
                
        except Exception as e:
            logger.error(f"点击导览框失败 {element_id}: {e}")
    
    async def process_cell(self, title: str, value: Any):
        """处理单个单元格的内容"""
        logger.info(f"进入process_cell方法: title={title}, value={value}")
        
        if pd.isna(value) or value == "":
            logger.info(f"跳过空值: {title}")
            return
            
        element_id = self.get_object_id(title)
        logger.info(f"获取到的element_id: {element_id}")
        
        if not element_id:
            logger.warning(f"未找到标题 '{title}' 对应的ID映射")
            return
            
        value_str = str(value).strip()
        logger.info(f"处理值: {value_str}")
        
        # 特殊处理：网上预约报账按钮（优先处理）
        if title == "网上预约报账按钮":
            logger.info(f"特殊处理网上预约报账按钮: {element_id}")
            # 从element_id中提取WF_YB6参数
            if "navToPrj('WF_YB6')" in element_id:
                logger.info("检测到navToPrj函数，调用click_navigation_panel")
                await self.click_navigation_panel("", "WF_YB6")
                return
            else:
                logger.info("未检测到navToPrj函数，尝试直接点击")
                # 如果element_id不是JavaScript函数，尝试直接点击
                await self.click_button(element_id)
                return
        
        # 处理按钮点击操作（以$开头）
        if value_str.startswith('$'):
            logger.info(f"处理按钮点击操作: {element_id}")
            await self.click_button(element_id)
            return
        
        # 处理系统导览框点击操作（以@开头）
        if value_str.startswith('@'):
            logger.info(f"处理系统导览框操作: {element_id}")
            nav_value = value_str[1:]  # 去掉@符号
            await self.click_navigation_panel(element_id, nav_value)
            return
        
        # 处理普通输入框
        logger.info(f"处理普通输入框: {element_id}")
        await self.fill_input(element_id, value_str)
    
    async def handle_login_with_captcha(self, record_data: pd.DataFrame):
        """处理登录流程，包括验证码输入"""
        logger.info("开始处理登录流程...")
        
        # 填写工号
        if "登录界面工号" in record_data.columns:
            uid = record_data["登录界面工号"].iloc[0]
            if pd.notna(uid) and uid != "":
                await self.fill_input("uid", str(uid))
        
        # 填写密码
        if "登录界面密码" in record_data.columns:
            pwd = record_data["登录界面密码"].iloc[0]
            if pd.notna(pwd) and pwd != "":
                await self.fill_input("pwd", str(pwd))
        
        # 等待用户输入验证码
        logger.info("密码填写完成，请在终端输入验证码...")
        captcha = input("请输入验证码: ")
        
        # 查找验证码输入框并填写
        try:
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
                    await self.page.wait_for_selector(selector, timeout=2000)
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
                await self.click_button("zhLogin")
        
        # 等待登录完成
        logger.info("登录请求已发送，等待页面跳转...")
        await asyncio.sleep(5)
        
        # 登录完成后，继续处理当前记录中的其他操作
        logger.info("登录完成，继续处理当前记录中的其他操作...")
        await self.process_record_after_login(record_data)
    
    async def process_record_after_login(self, record_data: pd.DataFrame):
        """登录后处理当前记录中的其他操作"""
        # 处理当前记录中的所有列（除了登录相关列）
        row = record_data.iloc[0]
        for col in record_data.columns:
            # 跳过登录相关的列和序号列
            if col in ["序号", "登录界面工号", "登录界面密码", "登录按钮"]:
                continue
            
            value = row[col]
            if pd.notna(value) and value != "":
                logger.info(f"处理登录后的操作: {col} = {value}")
                await self.process_cell(col, value)
    
    async def run_test(self):
        """运行测试"""
        try:
            # 加载数据
            await self.load_data()
            
            # 启动浏览器
            async with async_playwright() as p:
                self.browser = await p.chromium.launch(headless=False)
                self.page = await self.browser.new_page()
                
                # 导航到目标页面
                await self.page.goto("https://cwcx.uestc.edu.cn/WFManager/home.jsp")
                logger.info("成功导航到页面")
                
                # 等待页面加载
                await asyncio.sleep(3)
                
                # 只处理第一条记录（包含登录信息）
                first_record = self.reimbursement_data.iloc[0:1]
                logger.info("开始处理第一条记录（登录流程）")
                await self.handle_login_with_captcha(first_record)
                
                logger.info("测试完成")
                
        except Exception as e:
            logger.error(f"测试失败: {e}")
            raise
        finally:
            if self.browser:
                await self.browser.close()

async def main():
    """主函数"""
    # 检查文件是否存在
    if not os.path.exists("报销信息.xlsx"):
        logger.error("报销信息文件不存在")
        return
    
    if not os.path.exists("标题-ID.xlsx"):
        logger.error("标题-ID映射文件不存在")
        return
    
    # 创建测试实例并运行
    test = TestLoginFlow()
    await test.run_test()

if __name__ == "__main__":
    asyncio.run(main()) 