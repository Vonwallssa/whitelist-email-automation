import os
import glob
import smtplib
import getpass
import pandas as pd
import re
import sys
import shutil
import time
from email.message import EmailMessage
import argparse

# 添加邮箱验证函数
def is_valid_email(email):
    """
    验证邮箱格式是否正确
    Args:
        email: 要验证的邮箱字符串
    Returns:
        bool: 邮箱格式是否有效
    """
    # 基本的邮箱格式正则表达式
    email_pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
    return bool(email_pattern.match(email))

# 添加 sanitize_header 函数，用于清洗 header 值中的换行符
def sanitize_header(value: str) -> str:
    """移除字符串中的 CR/LF，防止 header 验证错误"""
    return re.sub(r"[\r\n]+", " ", str(value)).strip()

def verify_email_agreement_match(excel_path, target_dir):
    """
    验证test.xlsx中的航司对接人邮箱和协议号与target目录中的文件一致性
    Verify the consistency between airline contact emails and agreement numbers
    """
    # 检查Excel文件是否存在
    if not os.path.exists(excel_path):
        print(f"错误: Excel文件不存在 - {excel_path}")
        return {}
        
    # 检查目标目录是否存在
    if not os.path.exists(target_dir):
        print(f"错误: 目标目录不存在 - {target_dir}")
        return {}
    
    # 读取Excel文件
    try:
        df = pd.read_excel(excel_path)
        print(f"成功读取文件: {excel_path}，包含 {len(df)} 行数据")
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return {}
    
    # 检查必要的列是否存在
    required_columns = ['航司对接人邮箱', '协议号']
    if not all(col in df.columns for col in required_columns):
        print(f"Excel文件缺少必要的列: {', '.join(required_columns)}")
        return {}
    
    # 按邮箱地址聚合验证结果
    email_results = {}
    invalid_emails_count = 0
    
    # 遍历每一行数据
    for idx, row in df.iterrows():
        email = str(row['航司对接人邮箱']).strip()
        agreement_id = str(row['协议号']).strip()
        
        if not email or pd.isna(email) or email == 'nan':
            print(f"第 {idx+2} 行: 航司对接人邮箱为空")
            continue
            
        # 验证主收件人邮箱格式
        if not is_valid_email(email):
            print(f"第 {idx+2} 行: 航司对接人邮箱 '{email}' 格式不正确，跳过")
            invalid_emails_count += 1
            continue
            
        if not agreement_id or pd.isna(agreement_id) or agreement_id == 'nan':
            print(f"第 {idx+2} 行: 协议号为空")
            continue
        
        # 获取抄送列表
        cc_str = ""
        valid_cc_emails = []
        if '抄送邮箱' in row and row['抄送邮箱'] and not pd.isna(row['抄送邮箱']):
            cc_str_original = str(row['抄送邮箱']).strip()
            # 分割多个抄送邮箱
            cc_emails = [email.strip() for email in re.split(r'[,;\r\n]+', cc_str_original) if email.strip()]
            # 验证每个抄送邮箱格式
            for cc_email in cc_emails:
                if is_valid_email(cc_email):
                    valid_cc_emails.append(cc_email)
                else:
                    print(f"第 {idx+2} 行: 抄送邮箱 '{cc_email}' 格式不正确，将被忽略")
            
            # 使用有效的抄送邮箱重建抄送字符串
            cc_str = ",".join(valid_cc_emails)
        
        # 检查是否单独发送
        is_send_separately = False
        if '是否单独发送' in row and row['是否单独发送'] and not pd.isna(row['是否单独发送']):
            is_send_separately = str(row['是否单独发送']).strip() == '是'
        
        # 检查邮箱对应的文件夹是否存在
        email_folder = os.path.join(target_dir, email)
        folder_exists = os.path.isdir(email_folder)
        
        # 检查协议号对应的Excel文件是否存在
        excel_files = []
        matching_files = []
        
        if folder_exists:
            # 获取该文件夹下所有Excel文件
            patterns = ["*.xls", "*.xlsx", "*.xlsm"]
            for pat in patterns:
                excel_files.extend(glob.glob(os.path.join(email_folder, pat)))
            
            # 检查是否有文件名包含协议号的Excel
            for file_path in excel_files:
                filename = os.path.basename(file_path)
                if agreement_id in filename:
                    matching_files.append(file_path)
        
        # 保存验证结果
        match_found = len(matching_files) > 0
        
        # 如果邮箱不存在于结果字典中，创建新条目
        if email not in email_results:
            email_results[email] = {
                'folder_exists': folder_exists,
                'groups': {}  # 按抄送列表分组
            }
        
        # 如果需要单独发送，使用协议号作为额外分组依据
        group_key = cc_str
        if is_send_separately and match_found:
            # 为每个文件创建单独的分组键
            for match_file in matching_files:
                file_group_key = f"{cc_str}_{os.path.basename(match_file)}"
                if file_group_key not in email_results[email]['groups']:
                    email_results[email]['groups'][file_group_key] = {
                        'matches': [match_file],  # 单独发送时，一个分组只包含一个文件
                        'match_found': True,
                        'all_excels': [match_file],
                        'row_data': row.to_dict(),
                        'is_send_separately': True
                    }
        else:
            # 使用抄送列表作为分组标识（不单独发送的情况）
            if group_key not in email_results[email]['groups']:
                email_results[email]['groups'][group_key] = {
                    'matches': [],
                    'match_found': False,
                    'all_excels': [],
                    'row_data': row.to_dict(),
                    'is_send_separately': False
                }
            
            # 保存此条协议的匹配结果
            if match_found:
                email_results[email]['groups'][group_key]['match_found'] = True
                email_results[email]['groups'][group_key]['matches'].extend(matching_files)
            
            # 所有Excel文件都保存，不管是否匹配
            if folder_exists:
                email_results[email]['groups'][group_key]['all_excels'].extend(excel_files)
        
        # 打印每个协议号的验证结果
        status = "通过" if folder_exists and match_found else "失败"
        separate_info = "（单独发送）" if is_send_separately else ""
        print(f"验证 {email} - {agreement_id}{separate_info}: {status}")
        if not folder_exists:
            print(f"  - 文件夹不存在: {email_folder}")
        elif not match_found:
            print(f"  - 未找到包含协议号 {agreement_id} 的Excel文件")
        else:
            print(f"  - 找到匹配文件: {[os.path.basename(f) for f in matching_files]}")
            if valid_cc_emails and len(valid_cc_emails) > 0:
                print(f"  - 有效抄送邮箱: {valid_cc_emails}")
    
    # 打印按邮箱聚合的验证结果摘要
    for email, result in email_results.items():
        if not result['folder_exists']:
            print(f"\n邮箱 {email}: 文件夹不存在")
            continue
            
        for group_key, group_data in result['groups'].items():
            matches_count = len(group_data['matches'])
            # 对于单独发送的邮件，分组键包含文件名
            if group_data.get('is_send_separately', False):
                cc_part = group_key.split('_')[0] if '_' in group_key else ''
                cc_display = cc_part if cc_part else "无抄送"
                print(f"\n邮箱 {email} (抄送: {cc_display}) (单独发送): 找到 {matches_count} 个匹配文件")
            else:
                cc_display = group_key if group_key else "无抄送"
                if matches_count > 0:
                    print(f"\n邮箱 {email} (抄送: {cc_display}): 找到 {matches_count} 个匹配文件")
                else:
                    print(f"\n邮箱 {email} (抄送: {cc_display}): 未找到匹配文件")
    
    # 打印邮箱格式验证结果
    if invalid_emails_count > 0:
        print(f"\n注意: 发现 {invalid_emails_count} 行数据包含格式不正确的邮箱地址，这些行已被跳过")
            
    return email_results

def move_sent_files(sent_files, target_dir):
    """
    将已成功发送的文件移动到'已批量发送'文件夹
    Move successfully sent files to '已批量发送' folder
    """
    # 创建目标文件夹路径
    sent_folder = os.path.join(target_dir, '已批量发送')
    
    # 如果目标文件夹不存在，创建它
    if not os.path.exists(sent_folder):
        try:
            os.makedirs(sent_folder)
            print(f"创建文件夹: {sent_folder}")
        except Exception as e:
            print(f"创建文件夹失败: {e}")
            return False
    
    # 移动文件
    success_count = 0
    failed_count = 0
    failed_files = []
    
    for file_path in sent_files:
        if os.path.exists(file_path):
            try:
                filename = os.path.basename(file_path)
                destination = os.path.join(sent_folder, filename)
                shutil.move(file_path, destination)
                print(f"移动文件: {filename} -> {sent_folder}")
                success_count += 1
            except Exception as e:
                print(f"移动文件失败 {file_path}: {e}")
                failed_count += 1
                failed_files.append((file_path, str(e)))
        else:
            print(f"文件不存在，无法移动: {file_path}")
            failed_count += 1
            failed_files.append((file_path, "文件不存在"))
    
    print(f"\n文件移动摘要: 成功 {success_count} 个，失败 {failed_count} 个")
    if failed_count > 0:
        print("失败详情:")
        for file_path, error in failed_files:
            print(f"  - {file_path}: {error}")
    return True

#发送延时
def send_customized_emails(smtp_host: str,
                           smtp_port: int,
                           sender: str,
                           password: str,
                           validation_results: dict,
                           target_dir: str,
                           test_mode=False,
                           delay_seconds=1):
    """
    根据验证结果发送定制化的邮件
    Send customized emails based on validation results
    
    Args:
        smtp_host: SMTP服务器地址
        smtp_port: SMTP服务器端口
        sender: 发件人邮箱
        password: 发件人密码
        validation_results: 验证结果
        target_dir: 目标目录
        test_mode: 是否为测试模式
        delay_seconds: 每封邮件发送后的延迟秒数
    """
    if not validation_results:
        print("没有有效的验证结果，无法发送邮件")
        return
    
    # 跟踪已成功发送的文件夹
    sent_folders = set()
    
    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            
            try:
                server.login(sender, password)
            except smtplib.SMTPAuthenticationError:
                print(f"SMTP认证失败，请检查邮箱 {sender} 和密码是否正确")
                return
            
            # 跟踪成功和失败的邮件
            success_count = 0
            failed_count = 0
            email_count = 0
            
            for recipient, result in validation_results.items():
                if not result['folder_exists']:
                    print(f"跳过 {recipient}: 文件夹不存在")
                    failed_count += 1
                    continue
                
                # 处理每个抄送分组
                for group_key, group_data in result['groups'].items():
                    if not group_data['match_found']:
                        # 为单独发送的邮件提取抄送信息
                        if group_data.get('is_send_separately', False):
                            cc_part = group_key.split('_')[0] if '_' in group_key else ''
                            cc_display = cc_part if cc_part else "无抄送"
                        else:
                            cc_display = group_key if group_key else "无抄送"
                        print(f"跳过 {recipient} (抄送: {cc_display}): 未找到匹配的附件")
                        failed_count += 1
                        continue
                    
                    # 如果不是第一封邮件，添加延迟
                    if email_count > 0 and not test_mode:
                        print(f"延迟 {delay_seconds} 秒后继续发送...")
                        time.sleep(delay_seconds)
                    
                    email_count += 1
                    
                    # 获取邮件内容
                    row_data = group_data['row_data']
                    
                    # 获取抄送列表
                    cc_list = []
                    
                    # 对于单独发送的邮件，从分组键提取抄送信息
                    if group_data.get('is_send_separately', False):
                        cc_part = group_key.split('_')[0] if '_' in group_key else ''
                        if cc_part:
                            # 支持多个抄送邮箱，用逗号、分号或换行分隔
                            # 使用正则分割：逗号, 分号; 换行 \r 或 \n
                            cc_list = [email.strip() for email in re.split(r'[,;\r\n]+', cc_part) if email.strip()]
                    else:
                        if group_key:  # group_key 就是 cc_str
                            # 支持多个抄送邮箱，用逗号、分号或换行分隔
                            # 使用正则分割：逗号, 分号; 换行 \r 或 \n
                            cc_list = [email.strip() for email in re.split(r'[,;\r\n]+', group_key) if email.strip()]
                    
                    # 获取所有Excel文件并准备邮件主题+正文
                    all_excels = group_data['matches']
                    # 生成邮件正文：根据用户模板将附件名称列出
                    attachment_names = [os.path.basename(p) for p in all_excels]
                    body_lines = []
                    body_lines.append("经理，您好")
                    body_lines.append("")
                    body_lines.append("附件为本期白名单新增，烦请录入，谢谢！")
                    # 列出所有附件文件名
                    for name in attachment_names:
                        body_lines.append(name)
                    body_lines.append("")
                    body_lines.append("祝好。")
                    body_lines.append("")
                    body_lines.append("姓名/Name：请替换为你的姓名")
                    body_lines.append("部门/Dept：请替换为你的部门")
                    body_lines.append("电话/Tel：请替换为你的电话")
                    body_lines.append("邮箱/Email：请替换为你的邮箱")
                    body_lines.append("官网/Web：请替换为你的官网")
                    body_lines.append("地址/Add：请替换为你的地址")
                    custom_body = "\n".join(body_lines)
                    
                    # 构造 HTML 邮件正文，设置字体 '微软雅黑'，字号 14px
                    html_lines = [f"<p style='font-family:Microsoft YaHei; font-size:14px; margin:0 0 10px 0;'>{line if line else '&nbsp;'}</p>" for line in body_lines]
                    html_body = f"<html><body>{''.join(html_lines)}</body></html>"
                    
                    # 构建主题：单附件时使用"附件名_白名单新增"，多附件时使用原有逻辑
                    if len(all_excels) == 1:
                        # 单个附件时，使用整个附件名（去除.xlsx后缀）
                        filename = os.path.basename(all_excels[0])
                        # 移除文件扩展名
                        filename_without_ext = os.path.splitext(filename)[0]
                        subject = f"{filename_without_ext}_白名单新增"
                    elif len(all_excels) > 1:
                        # 多个附件时，使用原有逻辑
                        first_file = os.path.basename(all_excels[0])
                        m = re.match(r'^([A-Z]{2})', first_file)
                        code = m.group(1) if m else ''
                        subject = f"{code}_白名单新增_{len(all_excels)}家"
                    else:
                        subject = "白名单新增_0家"
                    
                    # 构造邮件
                    msg = EmailMessage()
                    msg["From"] = sanitize_header(sender)
                    msg["To"] = sanitize_header(recipient)
                    msg["Subject"] = sanitize_header(subject)
                    if cc_list:
                        msg["Cc"] = sanitize_header(", ".join(cc_list))
                    # 设置邮件正文：纯文本和 HTML 两个版本
                    msg.set_content(custom_body)
                    msg.add_alternative(html_body, subtype='html')
                    
                    # 添加附件（已经在上面获取了文件列表）
                    for file_path in all_excels:
                        try:
                            with open(file_path, "rb") as f:
                                data = f.read()
                            filename = os.path.basename(file_path)
                            msg.add_attachment(data, maintype="application", subtype="octet-stream",
                                               filename=filename)
                            print(f"  - 添加附件: {filename}")
                        except Exception as e:
                            print(f"  添加附件 {file_path} 失败: {e}")
                    
                    # 发送邮件
                    to_addrs = [recipient] + cc_list
                    # 显示用的抄送信息
                    if group_data.get('is_send_separately', False):
                        cc_part = group_key.split('_')[0] if '_' in group_key else ''
                        cc_display = cc_part if cc_part else "无抄送"
                        separate_info = "（单独发送）"
                    else:
                        cc_display = group_key if group_key else "无抄送"
                        separate_info = ""
                        
                    if test_mode:
                        print(f"测试模式: 将发送邮件给 {recipient} (抄送: {cc_display}){separate_info}")
                        print(f"  附件数量: {len(all_excels)}")
                        print(f"  附件列表: {[os.path.basename(f) for f in all_excels]}")
                        print(f"  邮件主题: {subject}")
                        success_count += 1
                        continue
                    try:
                        server.send_message(msg, from_addr=sender, to_addrs=to_addrs)
                        # 打印实际发送的附件列表
                        print(f"发送成功:")
                        print(f"  - 收件人: {recipient}")
                        print(f"  - 抄送: {cc_list}")
                        print(f"  - 主题: {subject}")
                        print(f"  - 附件: {[os.path.basename(p) for p in all_excels]}{separate_info}")
                        success_count += 1
                        
                        # 将已成功发送的文件夹添加到集合
                        folder_path = os.path.dirname(all_excels[0])
                        sent_folders.add(folder_path)
                    except Exception as e:
                        print(f"发送失败 {recipient} (抄送: {cc_display}){separate_info}: {e}")
                        failed_count += 1
            
            print(f"\n邮件发送摘要: 成功 {success_count} 封，失败 {failed_count} 封")
    except Exception as e:
        print(f"连接SMTP服务器 {smtp_host}:{smtp_port} 失败: {e}")
    
    # 移动已成功发送的文件夹
    if sent_folders and not test_mode:
        print("\n开始移动已成功发送的文件夹...")
        move_sent_folders(sent_folders, target_dir)

def move_sent_folders(folders, target_dir):
    """
    将已成功发送的文件夹移动到'已批量发送'文件夹
    Move successfully sent folders to '已批量发送' folder
    """
    # 创建目标文件夹路径
    sent_folder = os.path.join(target_dir, '已批量发送')
    
    # 如果目标文件夹不存在，创建它
    if not os.path.exists(sent_folder):
        try:
            os.makedirs(sent_folder)
            print(f"创建文件夹: {sent_folder}")
        except Exception as e:
            print(f"创建文件夹失败: {e}")
            return False
    
    # 移动文件夹
    success_count = 0
    failed_count = 0
    failed_folders = []
    
    for folder_path in folders:
        if os.path.exists(folder_path):
            try:
                folder_name = os.path.basename(folder_path)
                destination = os.path.join(sent_folder, folder_name)
                
                # 如果目标文件夹已存在，先删除
                if os.path.exists(destination):
                    shutil.rmtree(destination)
                
                # 移动文件夹
                shutil.move(folder_path, destination)
                print(f"移动文件夹: {folder_name} -> {sent_folder}")
                success_count += 1
            except Exception as e:
                print(f"移动文件夹失败 {folder_path}: {e}")
                failed_count += 1
                failed_folders.append((folder_path, str(e)))
        else:
            print(f"文件夹不存在，无法移动: {folder_path}")
            failed_count += 1
            failed_folders.append((folder_path, "文件夹不存在"))
    
    print(f"\n文件夹移动摘要: 成功 {success_count} 个，失败 {failed_count} 个")
    if failed_count > 0:
        print("失败详情:")
        for folder_path, error in failed_folders:
            print(f"  - {folder_path}: {error}")
    return True

#发送延时
def main(test_mode=False, delay_seconds=1):
    """主函数，处理参数并执行邮件验证和发送"""
    # 配置参数 - 请替换为你实际的SMTP配置
    smtp_host = "请替换为你的SMTP服务器地址"
    smtp_port = 587  # 请替换为你的SMTP端口，一般为587或25
    sender = "请替换为你的发件人邮箱"
    password = "请替换为你的邮箱密码"  # 请替换为你的邮箱密码或应用专用密码
    
    # 定义路径 - 请替换为你实际的路径
    test_excel_path = r"请替换为你实际的路径\邮件批量发送\MU批量发送列表.xlsx"
    target_dir = r"请替换为你实际的路径\target"
    
    # 验证邮箱和协议号的匹配
    print("开始验证邮箱和协议号的匹配...")
    validation_results = verify_email_agreement_match(test_excel_path, target_dir)
    
    # 如果没有验证结果，则退出
    if not validation_results:
        print("验证失败，无法继续发送邮件")
        return
    
    # 打印验证结果摘要
    total_emails = len(validation_results)
    total_groups = sum(len(result['groups']) for result in validation_results.values())
    passed_groups = sum(sum(1 for group_data in result['groups'].values() if group_data['match_found']) 
                        for result in validation_results.values())
    
    print(f"\n验证结果摘要: 共 {total_emails} 个邮箱, {total_groups} 个邮件组合, 通过 {passed_groups} 个，失败 {total_groups - passed_groups} 个")
    
    if passed_groups == 0:
        print("没有通过验证的邮箱-协议号组合，无法发送邮件")
        return
    
    # ---- 预览邮件发送信息 ----
    print("\n---- 预览邮件发送信息 ----")
    for recipient, result in validation_results.items():
        if not result['folder_exists']:
            continue
            
        for group_key, group_data in result['groups'].items():
            if not group_data['match_found']:
                continue
                
            # 获取抄送列表
            cc_list = []
            
            # 对于单独发送的邮件，从分组键提取抄送信息
            if group_data.get('is_send_separately', False):
                cc_part = group_key.split('_')[0] if '_' in group_key else ''
                if cc_part:
                    cc_list = [e.strip() for e in re.split(r'[,;\r\n]+', cc_part) if e.strip()]
                separate_info = "（单独发送）"
            else:
                if group_key:  # group_key 就是 cc_str
                    cc_list = [e.strip() for e in re.split(r'[,;\r\n]+', group_key) if e.strip()]
                separate_info = ""
                
            # 获取附件列表
            all_excels = group_data['matches']
            
            # 生成正文预览所需附件名列表
            attachment_names = [os.path.basename(p) for p in all_excels]
            
            # 检测 Excel 文件名前缀是否一致
            prefixes = []
            for file_path in all_excels:
                fname = os.path.basename(file_path)
                m = re.match(r'^([A-Z]{2})', fname)
                prefixes.append(m.group(1) if m else None)
                
            if len(prefixes) > 0:
                unique_prefixes = set(prefixes)
                if None in unique_prefixes or len(unique_prefixes) > 1:
                    # 对于单独发送的邮件，从分组键提取抄送信息
                    if group_data.get('is_send_separately', False):
                        cc_part = group_key.split('_')[0] if '_' in group_key else ''
                        cc_display = cc_part if cc_part else "无抄送"
                    else:
                        cc_display = group_key if group_key else "无抄送"
                    print(f"错误: 邮箱 {recipient} (抄送: {cc_display}){separate_info} 的 Excel 文件名前缀不一致: {prefixes}")
                    return
                    
            # 构建主题
            if len(all_excels) == 1:
                # 单个附件时，使用整个附件名（去除.xlsx后缀）
                filename = os.path.basename(all_excels[0])
                # 移除文件扩展名
                filename_without_ext = os.path.splitext(filename)[0]
                subject = f"{filename_without_ext}_白名单新增"
            elif len(all_excels) > 1:
                # 多个附件时，使用原有逻辑
                first_file = os.path.basename(all_excels[0])
                m = re.match(r'^([A-Z]{2})', first_file)
                code = m.group(1) if m else ''
                subject = f"{code}_白名单新增_{len(all_excels)}家"
            else:
                subject = "白名单新增_0家"
                
            # 打印预览信息
            # 对于单独发送的邮件，从分组键提取抄送信息
            if group_data.get('is_send_separately', False):
                cc_part = group_key.split('_')[0] if '_' in group_key else ''
                cc_display = cc_part if cc_part else "无抄送"
            else:
                cc_display = group_key if group_key else "无抄送"
                
            print(f"\n收件人: {recipient} (抄送: {cc_display}){separate_info}")
            print(f"抄送: {cc_list}")
            print(f"主题: {subject}")
            print(f"附件数量: {len(all_excels)}，文件: {attachment_names}")
            
            # 生成正文预览
            preview_lines = []
            preview_lines.append("经理，您好")
            preview_lines.append("")
            preview_lines.append("附件为本期白名单新增，烦请录入，谢谢！")
            for name in attachment_names:
                preview_lines.append(name)
            preview_lines.append("")
            preview_lines.append("祝好。")
            preview_lines.append("")
            preview_lines.append("姓名/Name：请替换为你的姓名")
            preview_lines.append("部门/Dept：请替换为你的部门")
            preview_lines.append("电话/Tel：请替换为你的电话")
            preview_lines.append("邮箱/Email：请替换为你的邮箱")
            preview_lines.append("官网/Web：请替换为你的官网")
            preview_lines.append("地址/Add：请替换为你的地址")
            preview_body = "\n".join(preview_lines)
            print("正文预览:\n" + preview_body)
            print("----------------------------------------")
    print("---- 预览结束 ----\n")
    
    # 确认是否继续发送邮件
    proceed = input("是否继续发送邮件？(y/n): ").strip().lower()
    if proceed != 'y':
        print("操作已取消")
        return
    
    # 打印发送间隔信息
    if not test_mode:
        print(f"\n已设置每封邮件发送间隔为 {delay_seconds} 秒")
    
    # 发送邮件，传入目标目录
    print("开始发送邮件...")
    send_customized_emails(smtp_host, smtp_port, sender, password, validation_results, target_dir, test_mode, delay_seconds)

if __name__ == "__main__":
    # 创建参数解析器
    parser = argparse.ArgumentParser(description='发送白名单邮件')
    parser.add_argument('--test', action='store_true', help='测试模式：验证逻辑但不发送邮件')
    parser.add_argument('--delay', type=int, default=2, help='每封邮件发送后的延迟秒数，默认为2秒') #发送延时
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 运行主函数
    main(test_mode=args.test, delay_seconds=args.delay) 