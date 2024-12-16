import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import re
from email.utils import parsedate_to_datetime
import os
import sys
from config import EMAIL_ADDRESS, EMAIL_PASSWORD


def decode_str(s):
    """解码邮件主题或发件人"""
    try:
        decoded_list = decode_header(s)
        result = ""
        for decoded_str, charset in decoded_list:
            if isinstance(decoded_str, bytes):
                if charset:
                    result += decoded_str.decode(charset)
                else:
                    result += decoded_str.decode('utf-8', 'ignore')
            else:
                result += decoded_str
        return result
    except Exception as e:
        print(f"解码错误: {str(e)}")
        return s

def get_email_content(msg, debug=False):
    """获取邮件内容，提取链接和验证码"""
    content = ""
    original_html = ""
    results = []
    
    # 优先获取HTML内容
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                try:
                    original_html = part.get_payload(decode=True).decode('utf-8', 'ignore')
                except Exception:
                    try:
                        original_html = part.get_payload(decode=True).decode('gbk', 'ignore')
                    except Exception as e:
                        continue
    
    # 如果没有找到HTML内容，尝试直接获取
    if not original_html and msg.get_content_type() == "text/html":
        try:
            original_html = msg.get_payload(decode=True).decode('utf-8', 'ignore')
        except:
            try:
                original_html = msg.get_payload(decode=True).decode('gbk', 'ignore')
            except:
                pass

    # 从HTML中提取内容
    if original_html:
        # 清理HTML中的换行和多余空格
        original_html = re.sub(r'\s+', ' ', original_html)
        
        # 提取验证码
        code_patterns = [
            r'验证码[：:]\s*?(\d{4,6})',
            r'verification code[：:]\s*?(\d{4,6})',
            r'code[：:]\s*?[\"\']*(\d{4,6})',
            r'[\s>](\d{6})[\s<]',
            r'<b[^>]*>(\d{4,6})</b>',
            r'<label[^>]*>(\d{4,6})</label>'
        ]
        
        # 提取验证码
        for pattern in code_patterns:
            code_match = re.search(pattern, original_html, re.IGNORECASE)
            if code_match:
                results.append(f"验证码: {code_match.group(1)}")
                break
        
        # 提取主要链接（第一个有效链接）
        link_matches = list(re.finditer(r'<a[^>]*?href=[\'"]([^\'"]+)[\'"][^>]*>(.*?)</a>', original_html))
        
        for match in link_matches:
            link = match.group(1)
            text = re.sub(r'<[^>]+>', '', match.group(2))
            
            # 过滤掉不需要的链接
            if any(x in link.lower() for x in [
                '.jpg', '.jpeg', '.png', '.gif', '.webp', 'wf/open',
                'unsubscribe', 'privacy', 'help', '#', 'javascript:'
            ]):
                continue
                
            if not text.strip():
                continue
            
            # 添加链接和链接文本
            results.append(f"链接: {link}")
            results.append(f"链接文本: {text.strip()}")
            break  # 找到第一个有效链接后就退出

        if results:
            return "\n".join(results)

    return ""

def get_text_content(msg):
    """获取邮件的纯文本内容"""
    content = ""
    
    if msg.is_multipart():
        # 如果邮件包含多个部分，递归获取文本内容
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    text = part.get_payload(decode=True).decode('utf-8', 'ignore')
                    content += text + "\n"
                except:
                    try:
                        text = part.get_payload(decode=True).decode('gbk', 'ignore')
                        content += text + "\n"
                    except:
                        continue
    else:
        # 如果邮件是文本
        try:
            content = msg.get_payload(decode=True).decode('utf-8', 'ignore')
        except:
            try:
                content = msg.get_payload(decode=True).decode('gbk', 'ignore')
            except:
                content = "无法解码的内容"
    
    return content.strip()

def save_to_file(email_info):
    """保存邮件信息到文件"""
    # 创建输出目录
    output_dir = "email_results"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 生成文件名（使用当前时间）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(output_dir, f"email_results_{timestamp}.txt")
    
    # 写入文件
    with open(filename, "w", encoding="utf-8") as f:
        f.write(email_info)
    
    return filename

def fetch_emails(email_address, password, search_params, hours=24):
    all_email_info = []  # 存储所有邮件信息
    try:
        # 建立连接
        imap = imaplib.IMAP4_SSL('domain-imap.cuiqiu.com')
        imap.login(email_address, password)
        imap.select('INBOX')
        
        # 建搜索条件
        search_criteria = []
        since_date = (datetime.now() - timedelta(hours=hours))
        formatted_date = since_date.strftime("%d-%b-%Y")
        search_criteria.append(f'SINCE "{formatted_date}"')
        
        # 添加其他索条件...
        if search_params.get('sender'):
            search_criteria.append(f'FROM "{search_params["sender"]}"')
        if search_params.get('subject'):
            search_criteria.append(f'SUBJECT "{search_params["subject"]}"')
        # 添加收件人搜索条件
        if search_params.get('recipient'):
            search_criteria.append(f'TO "{search_params["recipient"]}"')
            
        # 添加调试信息
        print("\n搜索参数:")
        for key, value in search_params.items():
            print(f"{key}: {value}")
            
        search_string = '(' + ' '.join(search_criteria) + ')'
        print(f"IMAP搜索字符串: {search_string}")
        
        status, message_numbers = imap.search(None, search_string)
        print(f"搜索状态: {status}")
        print(f"找到的邮件数量: {len(message_numbers[0].split()) if message_numbers[0] else 0}")
        
        if status != 'OK' or not message_numbers[0]:
            print(f"\n在过去{hours}小时没有找到符合条件的件")
            return
        
        # 批量获取邮
        email_list = []
        message_ids = message_numbers[0].split()
        
        print(f"\n找到 {len(message_ids)} 封符合条件的邮件")
        
        # 每次处理多封邮件
        batch_size = 10
        for i in range(0, len(message_ids), batch_size):
            batch_ids = message_ids[i:i + batch_size]
            for num in batch_ids:
                try:
                    status, msg_data = imap.fetch(num, '(RFC822)')
                    if status != 'OK':
                        continue
                    
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)
                    
                    # 先获取邮件内容
                    content = get_email_content(email_message)
                    
                    # 只有当找到链接时才输出邮件信息
                    if content:
                        subject = decode_str(email_message.get('Subject', '无主题'))
                        sender = decode_str(email_message.get('From', '未知发件人'))
                        recipient = decode_str(email_message.get('To', '未知收件人'))
                        
                        date_str = email_message.get('Date', '')
                        if date_str:
                            try:
                                date_obj = parsedate_to_datetime(date_str)
                                local_date = date_obj.astimezone()
                                date = local_date.strftime("%Y-%m-%d %H:%M:%S %Z")
                            except Exception:
                                date = date_str
                        else:
                            date = '未知时间'
                        
                        # 构建邮件信息字符串
                        email_info = (
                            f"\n{'='*50}\n"
                            f"发件人: {sender}\n"
                            f"收件人: {recipient}\n"
                            f"主题: {subject}\n"
                            f"时间: {date}\n"
                            f"{'='*50}\n"
                            f"{content}\n"
                            f"{'='*50}\n"
                        )
                        
                        # 同时打印到控制台和保存到列表
                        print(email_info)
                        all_email_info.append(email_info)
                    
                except Exception as e:
                    error_msg = f"处理邮件时出错: {str(e)}"
                    print(error_msg)
                    all_email_info.append(f"\n{error_msg}\n")
                    continue
        
        # 如果有邮件信息，保存到文件
        if all_email_info:
            full_content = "\n".join(all_email_info)
            filename = save_to_file(full_content)
            print(f"\n邮件信息已保存到文件: {filename}")
        else:
            print(f"\n在过去{hours}小时没有找到符合条件的邮件")
            
    finally:
        try:
            imap.close()
            imap.logout()
        except:
            pass


def get_config():
    """获取配置信息"""
    try:
        from config import EMAIL_ADDRESS, EMAIL_PASSWORD
        return EMAIL_ADDRESS, EMAIL_PASSWORD
    except ImportError:
        print("未找到配置文件，请手动输入邮箱信息")
        email = input("请输入邮箱地址: ").strip()
        password = input("请输入邮箱密码: ").strip()
        return email, password

def main():
    try:
        print("欢迎使用邮件获取工具 (IMAP版本)")
        
        # 使用配置获取方式
        email_address, password = get_config()
        
        search_params = {}
        
        print("\n请选择搜索条件（可选）:")
        print("1. 按发件人搜索")
        print("2. 按主题关键词搜索")
        print("3. 按收件人搜索")
        print("4. 按邮件内容关键词搜索")
        choices = input("请输入选项序号（多个用逗号分隔，直接回车表示不使用任何过滤条件）: ").strip()
        
        if choices:
            for choice in choices.split(','):
                if choice.strip() == '1':
                    search_params['sender'] = input("请输入发件人邮箱: ").strip()
                elif choice.strip() == '2':
                    search_params['subject'] = input("请输入主题关键词: ").strip()
                elif choice.strip() == '3':
                    search_params['recipient'] = input("请输入收件人邮箱: ").strip()
                elif choice.strip() == '4':
                    search_params['content'] = input("请输入内容关键词: ").strip()
        
        hours = input("\n请输入要查找的时间范围(小时数，直接回车默认24小时): ")
        hours = int(hours) if hours.strip() else 24
        
        fetch_emails(email_address, password, search_params, hours)
        
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
    
    finally:
        print("\n按回车键退出程序...")
        input()

if __name__ == "__main__":
    main() 