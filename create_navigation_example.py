import pandas as pd
import os

def create_navigation_example():
    """创建包含系统导览框操作的示例文件"""
    
    # 示例数据 - 包含导览框操作
    data = {
        '序号': [1, 1, 1, 2, 2, 3],
        '姓名': ['张三', '张三', '张三', '李四', '李四', '王五'],
        '金额': [100, 200, 150, 300, 250, 180],
        '支付方式': ['现金', '银行卡', '支付宝', '现金', '银行卡', '微信'],
        '报销类型': ['差旅费', '餐饮费', '办公费', '差旅费', '餐饮费', '办公费'],
        '部门': ['技术部', '技术部', '技术部', '销售部', '销售部', '人事部'],
        '费用日期': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19', '2024-01-20'],
        '备注': ['出差交通费', '午餐费用', '办公用品', '出差住宿费', '晚餐费用', '办公用品'],
        '系统导航': ['@WF_YB6', '@WF_YB7', '@WF_YB8', '@WF_YB6', '@WF_YB7', '@WF_YB8'],  # 新增导览框操作
        '子序列开始': ['是', '', '', '是', '', ''],
        '子序列结束': ['', '', '是', '', '是', '']
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存到Excel文件
    with pd.ExcelWriter('报销信息_导览框示例.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='BaoXiao_sheet', index=False)
    
    print("成功创建包含导览框操作的示例Excel文件: 报销信息_导览框示例.xlsx")
    
    # 更新标题-ID映射文件
    mapping_data = {
        '标题': [
            '姓名', '金额', '支付方式', '报销类型', '部门', 
            '费用日期', '备注', '系统导航', '提交按钮', '保存按钮', '重置按钮'
        ],
        '网页元素ID': [
            'name_input', 'amount_input', 'payment_method', 'expense_type', 'department',
            'expense_date', 'remarks', 'nav_panel', 'submit_btn', 'save_btn', 'reset_btn'
        ]
    }
    
    # 创建DataFrame
    mapping_df = pd.DataFrame(mapping_data)
    
    # 保存到Excel文件
    mapping_df.to_excel('标题-ID_导览框示例.xlsx', index=False)
    
    print("成功创建包含导览框映射的标题-ID文件: 标题-ID_导览框示例.xlsx")
    
    print("\n使用说明:")
    print("1. 在Excel中，系统导航列使用 '@WF_YB6' 格式")
    print("2. @ 符号表示这是一个系统导览框点击操作")
    print("3. WF_YB6 是导览框的标识符，对应onclick='navToPrj(\"WF_YB6\")'")
    print("4. 程序会自动查找并点击对应的导览框元素")

if __name__ == "__main__":
    create_navigation_example() 