import pandas as pd

# 文件路径 - 请替换为你实际的路径
rawdata_path = r"请替换为你实际的路径\rawdata.xlsx"
contact_list_path = r"请替换为你实际的路径\contact_list.xlsx"
output_path = r"请替换为你实际的路径\output\whitelist_updated.xlsx"

# 读取文件
rawdata_df = pd.read_excel(rawdata_path)
contact_list_df = pd.read_excel(contact_list_path)

# 检查和替换
# 对协议号建立映射关系 {协议号: 协议客户名称}
protocol_mapping = dict(zip(contact_list_df['协议号'], contact_list_df['协议客户名称']))  # 使用正确的列名

# 替换公司名称
rawdata_df['公司名称'] = rawdata_df['协议号'].map(protocol_mapping).combine_first(rawdata_df['公司名称'])

# 保存修改后的文件
rawdata_df.to_excel(output_path, index=False)

print(f"文件已更新并保存到：{output_path}")
