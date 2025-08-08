import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def check_excel_files():
    """检查Excel文件内容"""
    try:
        # 检查标题-ID映射文件
        logger.info("检查标题-ID.xlsx文件...")
        title_id_df = pd.read_excel("标题-ID.xlsx")
        logger.info(f"标题-ID.xlsx 列名: {list(title_id_df.columns)}")
        logger.info(f"标题-ID.xlsx 数据:\n{title_id_df}")
        
        # 检查报销信息文件
        logger.info("\n检查报销信息.xlsx文件...")
        reimbursement_df = pd.read_excel("报销信息.xlsx", sheet_name="BaoXiao_sheet")
        logger.info(f"报销信息.xlsx 列名: {list(reimbursement_df.columns)}")
        logger.info(f"报销信息.xlsx 数据:\n{reimbursement_df}")
        
    except Exception as e:
        logger.error(f"检查Excel文件时发生错误: {e}")

if __name__ == "__main__":
    check_excel_files() 