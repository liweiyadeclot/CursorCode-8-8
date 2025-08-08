import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import logging
import os

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def test_automation():
    """测试自动化程序"""
    try:
        # 检查文件是否存在
        if not os.path.exists("报销信息.xlsx"):
            logger.error("报销信息文件不存在")
            return
        
        if not os.path.exists("标题-ID.xlsx"):
            logger.error("标题-ID映射文件不存在")
            return
        
        # 加载数据
        logger.info("开始加载数据...")
        
        # 加载标题-ID映射
        mapping_df = pd.read_excel("标题-ID.xlsx")
        title_id_mapping = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))
        logger.info(f"成功加载标题-ID映射，共{len(title_id_mapping)}条记录")
        
        # 加载报销信息数据
        reimbursement_data = pd.read_excel("报销信息.xlsx", sheet_name="BaoXiao_sheet")
        logger.info(f"成功加载报销信息数据，共{len(reimbursement_data)}行")
        
        # 启动浏览器
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            page = await browser.new_page()
            
            # 导航到测试页面
            test_url = "file:///C:/Users/FH/PycharmProjects/CursorCode8-5/test_page.html"
            await page.goto(test_url)
            logger.info(f"成功导航到测试页面: {test_url}")
            
            # 等待页面加载
            await asyncio.sleep(2)
            
            # 测试填写表单
            logger.info("开始测试表单填写...")
            
            # 获取第一条记录
            first_record = reimbursement_data.iloc[0]
            
            # 填写姓名
            if '姓名' in title_id_mapping and pd.notna(first_record['姓名']):
                name_id = title_id_mapping['姓名']
                await page.fill(f"#{name_id}", str(first_record['姓名']))
                logger.info(f"填写姓名: {first_record['姓名']}")
                await asyncio.sleep(0.5)
            
            # 填写金额
            if '金额' in title_id_mapping and pd.notna(first_record['金额']):
                amount_id = title_id_mapping['金额']
                await page.fill(f"#{amount_id}", str(first_record['金额']))
                logger.info(f"填写金额: {first_record['金额']}")
                await asyncio.sleep(0.5)
            
            # 选择支付方式
            if '支付方式' in title_id_mapping and pd.notna(first_record['支付方式']):
                payment_id = title_id_mapping['支付方式']
                await page.select_option(f"#{payment_id}", str(first_record['支付方式']))
                logger.info(f"选择支付方式: {first_record['支付方式']}")
                await asyncio.sleep(0.5)
            
            # 选择报销类型
            if '报销类型' in title_id_mapping and pd.notna(first_record['报销类型']):
                expense_id = title_id_mapping['报销类型']
                await page.select_option(f"#{expense_id}", str(first_record['报销类型']))
                logger.info(f"选择报销类型: {first_record['报销类型']}")
                await asyncio.sleep(0.5)
            
            # 选择部门
            if '部门' in title_id_mapping and pd.notna(first_record['部门']):
                dept_id = title_id_mapping['部门']
                await page.select_option(f"#{dept_id}", str(first_record['部门']))
                logger.info(f"选择部门: {first_record['部门']}")
                await asyncio.sleep(0.5)
            
            # 填写费用日期
            if '费用日期' in title_id_mapping and pd.notna(first_record['费用日期']):
                date_id = title_id_mapping['费用日期']
                await page.fill(f"#{date_id}", str(first_record['费用日期']))
                logger.info(f"填写费用日期: {first_record['费用日期']}")
                await asyncio.sleep(0.5)
            
            # 填写备注
            if '备注' in title_id_mapping and pd.notna(first_record['备注']):
                remarks_id = title_id_mapping['备注']
                await page.fill(f"#{remarks_id}", str(first_record['备注']))
                logger.info(f"填写备注: {first_record['备注']}")
                await asyncio.sleep(0.5)
            
            logger.info("表单填写完成，等待5秒后关闭浏览器...")
            await asyncio.sleep(5)
            
            await browser.close()
            logger.info("测试完成！")
            
    except Exception as e:
        logger.error(f"测试过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_automation()) 