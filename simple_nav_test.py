import asyncio
from playwright.async_api import async_playwright
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def simple_nav_test():
    """简化的导航测试"""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        
        try:
            # 导航到页面
            logger.info("导航到页面...")
            await page.goto("https://cwcx.uestc.edu.cn/WFManager/home.jsp")
            await asyncio.sleep(2)
            
            # 填写登录信息
            logger.info("填写登录信息...")
            await page.fill("input[name='uid']", "test_user")
            await page.fill("input[name='pwd']", "test_password")
            
            # 等待用户输入验证码
            captcha = input("请输入验证码: ")
            await page.fill("input[name='captcha']", captcha)
            
            # 点击登录
            logger.info("点击登录按钮...")
            await page.click("input[btnName='zhLogin']")
            await asyncio.sleep(3)
            
            # 尝试点击导航窗格
            logger.info("尝试点击导航窗格...")
            
            # 方法1: 直接点击第一个导航元素
            try:
                await page.click("#sysNavigator > div:first-child")
                logger.info("成功点击第一个导航元素")
            except Exception as e:
                logger.warning(f"方法1失败: {e}")
            
            # 方法2: 通过JavaScript执行
            try:
                await page.evaluate("navToPrj('WF_YB6')")
                logger.info("成功执行JavaScript导航")
            except Exception as e:
                logger.warning(f"方法2失败: {e}")
            
            # 等待观察结果
            logger.info("等待5秒观察结果...")
            await asyncio.sleep(5)
            
        except Exception as e:
            logger.error(f"测试失败: {e}")
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(simple_nav_test()) 