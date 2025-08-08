import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import logging
from config import *

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class NavigationTest:
    def __init__(self):
        self.page = None
        self.browser = None
        
    async def setup(self):
        """初始化浏览器"""
        playwright = await async_playwright().start()
        self.browser = await playwright.chromium.launch(headless=False)
        self.page = await self.browser.new_page()
        
    async def test_navigation_click(self):
        """测试导航窗格点击功能"""
        try:
            # 导航到目标页面
            await self.page.goto(TARGET_URL)
            logger.info(f"成功导航到页面: {TARGET_URL}")
            await asyncio.sleep(3)
            
            # 等待页面加载
            await self.page.wait_for_load_state('networkidle')
            
            # 测试多种点击方法
            value = "WF_YB6"
            
            # 方法1: 通过精确的class和onclick组合查找
            try:
                selector = f"div.syslink.ieSensitive.noie[onclick*='{value}']"
                logger.info(f"尝试方法1: {selector}")
                await self.page.wait_for_selector(selector, timeout=5000)
                await self.page.click(selector)
                logger.info("方法1成功: 通过精确class和onclick点击")
                await asyncio.sleep(2)
                return
            except Exception as e:
                logger.warning(f"方法1失败: {e}")
            
            # 方法2: 通过sysNavigator容器内的第一个元素查找
            try:
                selector = "#sysNavigator > div.syslink.ieSensitive.noie:first-child"
                logger.info(f"尝试方法2: {selector}")
                await self.page.wait_for_selector(selector, timeout=5000)
                await self.page.click(selector)
                logger.info("方法2成功: 通过sysNavigator第一个元素点击")
                await asyncio.sleep(2)
                return
            except Exception as e:
                logger.warning(f"方法2失败: {e}")
            
            # 方法3: 通过JavaScript直接执行navToPrj函数
            try:
                logger.info(f"尝试方法3: 执行navToPrj('{value}')")
                await self.page.evaluate(f"navToPrj('{value}')")
                logger.info("方法3成功: 通过JavaScript执行navToPrj函数")
                await asyncio.sleep(2)
                return
            except Exception as e:
                logger.warning(f"方法3失败: {e}")
            
            # 方法4: 通过JavaScript直接调用onclick事件
            try:
                script = f"""
                const elements = document.querySelectorAll('div[onclick*="{value}"]');
                if (elements.length > 0) {{
                    elements[0].onclick();
                    return true;
                }}
                return false;
                """
                result = await self.page.evaluate(script)
                if result:
                    logger.info("方法4成功: 通过JavaScript调用onclick事件")
                    await asyncio.sleep(2)
                    return
                else:
                    logger.warning("方法4失败: 未找到包含onclick的元素")
            except Exception as e:
                logger.warning(f"方法4失败: {e}")
            
            logger.error("所有方法都失败了")
            
        except Exception as e:
            logger.error(f"测试过程中发生错误: {e}")
    
    async def cleanup(self):
        """清理资源"""
        if self.browser:
            await self.browser.close()

async def main():
    test = NavigationTest()
    try:
        await test.setup()
        await test.test_navigation_click()
        # 等待一段时间观察结果
        await asyncio.sleep(5)
    finally:
        await test.cleanup()

if __name__ == "__main__":
    asyncio.run(main()) 