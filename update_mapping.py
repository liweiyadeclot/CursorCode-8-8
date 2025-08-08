import pandas as pd

# 读取现有的映射文件
mapping_df = pd.read_excel("标题-ID.xlsx")
print("当前映射内容:")
print(mapping_df)

# 检查是否已经存在"申请报销单按钮"的映射
if "申请报销单按钮" not in mapping_df.iloc[:, 0].values:
    # 添加新的映射
    new_row = pd.DataFrame([["申请报销单按钮", "申请报销单"]], columns=mapping_df.columns)
    mapping_df = pd.concat([mapping_df, new_row], ignore_index=True)
    print("\n添加了新的映射: 申请报销单按钮 -> 申请报销单")
else:
    # 更新现有映射
    mask = mapping_df.iloc[:, 0] == "申请报销单按钮"
    mapping_df.loc[mask, mapping_df.columns[1]] = "申请报销单"
    print("\n更新了现有映射: 申请报销单按钮 -> 申请报销单")

# 保存更新后的文件
mapping_df.to_excel("标题-ID.xlsx", index=False)
print("\n更新后的映射内容:")
print(mapping_df)
print("\n文件已保存!") 