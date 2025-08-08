import pandas as pd

def read_excel_files():
    """读取Excel文件内容"""
    
    try:
        # 读取标题-ID映射文件
        print("=== 标题-ID.xlsx 内容 ===")
        mapping_df = pd.read_excel("标题-ID.xlsx")
        print(mapping_df)
        print("\n")
        
        # 读取报销信息文件
        print("=== 报销信息.xlsx 内容 ===")
        reimbursement_df = pd.read_excel("报销信息.xlsx", sheet_name="BaoXiao_sheet")
        print(reimbursement_df)
        print("\n")
        
        # 显示列名
        print("=== 报销信息.xlsx 列名 ===")
        print(list(reimbursement_df.columns))
        
    except Exception as e:
        print(f"读取文件失败: {e}")

if __name__ == "__main__":
    read_excel_files() 