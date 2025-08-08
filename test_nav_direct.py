import asyncio
from playwright.async_api import async_playwright
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def test_navigation():
    """先登录，然后测试导航窗格点击"""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        
        try:
            # 导航到页面
            await page.goto("https://cwcx.uestc.edu.cn/WFManager/home.jsp")
            logger.info("成功导航到页面")
            await asyncio.sleep(3)
            
            # 等待页面加载
            await page.wait_for_load_state('networkidle')
            
            # 先进行登录
            logger.info("开始登录流程...")
            
            # 填写工号
            try:
                await page.fill("input[name='uid']", "test_user")
                logger.info("填写工号成功")
            except Exception as e:
                logger.warning(f"填写工号失败: {e}")
            
            # 填写密码
            try:
                await page.fill("input[name='pwd']", "test_password")
                logger.info("填写密码成功")
            except Exception as e:
                logger.warning(f"填写密码失败: {e}")
            
            # 等待用户输入验证码
            captcha = input("请输入验证码: ")
            logger.info("用户输入验证码完成")
            
            # 填写验证码
            try:
                await page.fill("input[name='captcha']", captcha)
                logger.info("填写验证码成功")
            except Exception as e:
                logger.warning(f"填写验证码失败: {e}")
            
            # 点击登录按钮
            try:
                await page.click("input[btnName='zhLogin']")
                logger.info("点击登录按钮成功")
                await asyncio.sleep(3)
            except Exception as e:
                logger.warning(f"点击登录按钮失败: {e}")
            
            # 等待登录后的页面加载
            await page.wait_for_load_state('networkidle')
            logger.info("登录后页面加载完成")
            
            # 现在测试导航窗格点击
            logger.info("开始测试导航窗格点击...")
            
            # 方法1: 直接通过JavaScript执行navToPrj函数
            logger.info("尝试方法1: 直接执行navToPrj('WF_YB6')")
            try:
                await page.evaluate("navToPrj('WF_YB6')")
                logger.info("方法1成功: 直接执行JavaScript函数")
                await asyncio.sleep(3)
            except Exception as e:
                logger.warning(f"方法1失败: {e}")
            
            # 方法2: 查找并点击第一个导航元素
            logger.info("尝试方法2: 点击第一个导航元素")
            try:
                # 等待sysNavigator容器出现
                await page.wait_for_selector("#sysNavigator", timeout=5000)
                
                # 点击第一个子元素
                await page.click("#sysNavigator > div:first-child")
                logger.info("方法2成功: 点击第一个导航元素")
                await asyncio.sleep(3)
            except Exception as e:
                logger.warning(f"方法2失败: {e}")
            
            # 方法3: 通过onclick属性查找
            logger.info("尝试方法3: 通过onclick属性查找")
            try:
                selector = "div[onclick*='navToPrj'][onclick*='WF_YB6']"
                await page.wait_for_selector(selector, timeout=5000)
                await page.click(selector)
                logger.info("方法3成功: 通过onclick属性点击")
                await asyncio.sleep(3)
            except Exception as e:
                logger.warning(f"方法3失败: {e}")
            
            # 方法4: 通过class和title组合查找
            logger.info("尝试方法4: 通过class和title组合查找")
            try:
                selector = "div.syslink[title='点击进入']"
                elements = await page.query_selector_all(selector)
                if elements:
                    await elements[0].click()
                    logger.info("方法4成功: 通过class和title点击第一个元素")
                    await asyncio.sleep(3)
                else:
                    logger.warning("方法4失败: 未找到元素")
            except Exception as e:
                logger.warning(f"方法4失败: {e}")
            
            # 等待一段时间观察结果
            logger.info("等待10秒观察结果...")
            await asyncio.sleep(10)
            
        except Exception as e:
            logger.error(f"测试过程中发生错误: {e}")
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(test_navigation()) 