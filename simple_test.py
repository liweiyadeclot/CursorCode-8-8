import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import logging
import os

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def simple_test():
    """简单的测试程序"""
    try:
        logger.info("开始简单测试...")
        
        # 检查文件是否存在
        if not os.path.exists("报销信息.xlsx"):
            logger.error("报销信息文件不存在")
            return
        
        if not os.path.exists("标题-ID.xlsx"):
            logger.error("标题-ID映射文件不存在")
            return
        
        logger.info("文件检查通过")
        
        # 加载数据
        logger.info("开始加载数据...")
        
        # 加载标题-ID映射
        mapping_df = pd.read_excel("标题-ID.xlsx")
        title_id_mapping = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))
        logger.info(f"成功加载标题-ID映射，共{len(title_id_mapping)}条记录")
        
        # 加载报销信息数据
        reimbursement_data = pd.read_excel("报销信息.xlsx", sheet_name="BaoXiao_sheet")
        logger.info(f"成功加载报销信息数据，共{len(reimbursement_data)}行")
        
        logger.info("数据加载完成")
        
        # 启动浏览器
        logger.info("启动浏览器...")
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            logger.info("浏览器启动成功")
            
            page = await browser.new_page()
            logger.info("新页面创建成功")
            
            # 导航到测试页面
            test_url = "file:///C:/Users/FH/PycharmProjects/CursorCode8-5/test_page.html"
            logger.info(f"准备导航到: {test_url}")
            
            await page.goto(test_url)
            logger.info("页面导航成功")
            
            # 等待页面加载
            await asyncio.sleep(2)
            logger.info("页面加载等待完成")
            
            # 简单的表单填写测试
            logger.info("开始表单填写测试...")
            
            # 获取第一条记录
            first_record = reimbursement_data.iloc[0]
            logger.info(f"第一条记录: {first_record.to_dict()}")
            
            # 测试填写姓名
            if '姓名' in title_id_mapping and pd.notna(first_record['姓名']):
                name_id = title_id_mapping['姓名']
                logger.info(f"准备填写姓名，ID: {name_id}, 值: {first_record['姓名']}")
                
                try:
                    await page.fill(f"#{name_id}", str(first_record['姓名']))
                    logger.info("姓名填写成功")
                except Exception as e:
                    logger.error(f"姓名填写失败: {e}")
            
            await asyncio.sleep(1)
            
            # 测试填写金额
            if '金额' in title_id_mapping and pd.notna(first_record['金额']):
                amount_id = title_id_mapping['金额']
                logger.info(f"准备填写金额，ID: {amount_id}, 值: {first_record['金额']}")
                
                try:
                    await page.fill(f"#{amount_id}", str(first_record['金额']))
                    logger.info("金额填写成功")
                except Exception as e:
                    logger.error(f"金额填写失败: {e}")
            
            await asyncio.sleep(1)
            
            logger.info("表单填写测试完成，等待3秒后关闭浏览器...")
            await asyncio.sleep(3)
            
            await browser.close()
            logger.info("浏览器已关闭，测试完成！")
            
    except Exception as e:
        logger.error(f"测试过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(simple_test()) 