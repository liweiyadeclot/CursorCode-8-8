import pandas as pd
import os

def create_sample_reimbursement_excel():
    """创建示例报销信息Excel文件"""
    
    # 示例数据
    data = {
        '序号': [1, 1, 1, 2, 2, 3],
        '姓名': ['张三', '张三', '张三', '李四', '李四', '王五'],
        '金额': [100, 200, 150, 300, 250, 180],
        '支付方式': ['现金', '银行卡', '支付宝', '现金', '银行卡', '微信'],
        '报销类型': ['差旅费', '餐饮费', '办公费', '差旅费', '餐饮费', '办公费'],
        '部门': ['技术部', '技术部', '技术部', '销售部', '销售部', '人事部'],
        '费用日期': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19', '2024-01-20'],
        '备注': ['出差交通费', '午餐费用', '办公用品', '出差住宿费', '晚餐费用', '办公用品'],
        '子序列开始': ['是', '', '', '是', '', ''],
        '子序列结束': ['', '', '是', '', '是', '']
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存到Excel文件
    with pd.ExcelWriter('报销信息.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='BaoXiao_sheet', index=False)
        
        # 创建其他sheet的示例
        df2 = df.copy()
        df2 = df2.head(5)  # 只取前5行作为另一个sheet的示例
        df2['序号'] = [4, 4, 5, 5, 6]  # 重新设置序号
        df2.to_excel(writer, sheet_name='Other_sheet', index=False)
    
    print("成功创建示例报销信息Excel文件: 报销信息.xlsx")

def create_sample_mapping_excel():
    """创建示例标题-ID映射Excel文件"""
    
    # 示例映射数据
    mapping_data = {
        '标题': [
            '姓名', '金额', '支付方式', '报销类型', '部门', 
            '费用日期', '备注', '提交按钮', '保存按钮', '重置按钮'
        ],
        '网页元素ID': [
            'name_input', 'amount_input', 'payment_method', 'expense_type', 'department',
            'expense_date', 'remarks', 'submit_btn', 'save_btn', 'reset_btn'
        ]
    }
    
    # 创建DataFrame
    df = pd.DataFrame(mapping_data)
    
    # 保存到Excel文件
    df.to_excel('标题-ID.xlsx', index=False)
    
    print("成功创建示例标题-ID映射Excel文件: 标题-ID.xlsx")

def create_sample_html():
    """创建示例HTML页面用于测试"""
    
    html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>报销信息填写系统</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .form-container {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #333;
        }
        input, select, textarea {
            width: 100%;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
        }
        .button-group {
            margin-top: 30px;
            text-align: center;
        }
        button {
            padding: 12px 24px;
            margin: 0 10px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        .submit-btn {
            background-color: #007bff;
            color: white;
        }
        .save-btn {
            background-color: #28a745;
            color: white;
        }
        .reset-btn {
            background-color: #dc3545;
            color: white;
        }
        button:hover {
            opacity: 0.8;
        }
    </style>
</head>
<body>
    <div class="form-container">
        <h1>报销信息填写</h1>
        <form id="reimbursement-form">
            <div class="form-group">
                <label for="name_input">姓名:</label>
                <input type="text" id="name_input" name="name" required>
            </div>
            
            <div class="form-group">
                <label for="amount_input">金额:</label>
                <input type="number" id="amount_input" name="amount" step="0.01" required>
            </div>
            
            <div class="form-group">
                <label for="payment_method">支付方式:</label>
                <select id="payment_method" name="payment_method" required>
                    <option value="">请选择支付方式</option>
                    <option value="现金">现金</option>
                    <option value="银行卡">银行卡</option>
                    <option value="支付宝">支付宝</option>
                    <option value="微信">微信</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="expense_type">报销类型:</label>
                <select id="expense_type" name="expense_type" required>
                    <option value="">请选择报销类型</option>
                    <option value="差旅费">差旅费</option>
                    <option value="餐饮费">餐饮费</option>
                    <option value="办公费">办公费</option>
                    <option value="交通费">交通费</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="department">部门:</label>
                <select id="department" name="department" required>
                    <option value="">请选择部门</option>
                    <option value="技术部">技术部</option>
                    <option value="销售部">销售部</option>
                    <option value="人事部">人事部</option>
                    <option value="财务部">财务部</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="expense_date">费用日期:</label>
                <input type="date" id="expense_date" name="expense_date" required>
            </div>
            
            <div class="form-group">
                <label for="remarks">备注:</label>
                <textarea id="remarks" name="remarks" rows="3"></textarea>
            </div>
            
            <div class="button-group">
                <button type="button" id="save_btn" class="save-btn">保存</button>
                <button type="button" id="reset_btn" class="reset-btn">重置</button>
                <button type="submit" id="submit_btn" class="submit-btn">提交</button>
            </div>
        </form>
    </div>
    
    <script>
        // 简单的表单处理
        document.getElementById('reimbursement-form').addEventListener('submit', function(e) {
            e.preventDefault();
            alert('表单提交成功！');
        });
        
        document.getElementById('save_btn').addEventListener('click', function() {
            alert('保存成功！');
        });
        
        document.getElementById('reset_btn').addEventListener('click', function() {
            document.getElementById('reimbursement-form').reset();
            alert('表单已重置！');
        });
    </script>
</body>
</html>
    """
    
    with open('test_page.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("成功创建示例HTML测试页面: test_page.html")

def main():
    """主函数"""
    print("开始创建示例文件...")
    
    # 创建示例文件
    create_sample_reimbursement_excel()
    create_sample_mapping_excel()
    create_sample_html()
    
    print("\n所有示例文件创建完成！")
    print("\n文件说明:")
    print("1. 报销信息.xlsx - 包含报销数据的Excel文件")
    print("2. 标题-ID.xlsx - 包含标题与网页元素ID映射的Excel文件")
    print("3. test_page.html - 用于测试的HTML页面")
    print("\n使用方法:")
    print("1. 修改 config.py 中的 TARGET_URL 为 'file:///path/to/test_page.html'")
    print("2. 运行 python reimbursement_automation.py")

if __name__ == "__main__":
    main() 