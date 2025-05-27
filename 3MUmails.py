import os
import shutil
import pandas as pd
import numpy as np

# 定义Excel文件路径 - 请替换为你实际的路径
mapping_file_path = r"请替换为你实际的路径\邮件批量发送\MU批量发送列表.xlsx"

# 读取映射文件，并确保协议号列被转换为字符串
mapping_df = pd.read_excel(mapping_file_path, dtype={'协议号': str, '航司对接人邮箱': str})

# 将映射关系存储在字典中，并处理 NaN 值
mapping_df['航司对接人邮箱'].fillna('无邮箱', inplace=True)
mapping = dict(zip(mapping_df['协议号'], mapping_df['航司对接人邮箱']))

# 打印映射关系以进行调试
print("映射关系：", mapping)

# 定义Excel文件所在的目录和目标根目录 - 请替换为你实际的路径
source_directory = r"请替换为你实际的路径\output"  # 请修改为实际路径
target_root_directory = r"请替换为你实际的路径\target"  # 请修改为实际路径

# 确保目标根目录存在
if not os.path.exists(target_root_directory):
    os.makedirs(target_root_directory)

# 遍历源目录中的文件
for filename in os.listdir(source_directory):
    if filename.endswith('.xlsx') and not filename.startswith('~$'):
        # 提取文件名中的编号，并转换为字符串
        parts = filename.split('_')
        if len(parts) > 1:
            number = parts[1]
            print(f"处理文件: {filename}, 提取到的编号: {number}")
            if number in mapping:
                # 获取对应的邮箱地址，并检查其有效性
                email = mapping[number].strip()  # 移除邮箱地址两端的空格
                if pd.notna(email):
                    print(f"编号 {number} 对应的邮箱地址是 {email}")
                    # 创建目标文件夹路径
                    target_directory = os.path.join(target_root_directory, email)
                    if not os.path.exists(target_directory):
                        os.makedirs(target_directory)
                        print(f"创建目标文件夹: {target_directory}")
                    # 检查文件是否存在于源目录
                    source_file_path = os.path.join(source_directory, filename)
                    target_file_path = os.path.join(target_directory, filename)
                    print(f"源文件路径: {source_file_path}")
                    print(f"目标文件路径: {target_file_path}")
                    if os.path.exists(source_file_path):
                        # 移动文件到目标文件夹
                        shutil.move(source_file_path, target_file_path)
                        print(f"移动文件 {filename} 到 {target_directory}")
                    else:
                        print(f"源文件 {filename} 不存在")
                else:
                    print(f"编号 {number} 对应的邮箱地址无效")
            else:
                print(f"编号 {number} 不在映射关系中")
        else:
            print(f"文件名 {filename} 格式不正确，无法提取编号")

print("文件移动完成。")
