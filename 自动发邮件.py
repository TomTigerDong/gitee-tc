#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
自动发送邮件程序
功能：发送带有附件的电子邮件
作者：[你的名字]
日期：[创建日期]
"""

import os
import logging
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.header import Header
from typing import Optional, List, Union
from pathlib import Path
from dotenv import load_dotenv

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_sender.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class EmailConfig:
    """邮件配置类"""
    def __init__(self):
        # 加载环境变量
        load_dotenv()
        
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.qq.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '465'))
        self.email = os.getenv('EMAIL')
        self.password = os.getenv('EMAIL_PASSWORD')
        
        if not all([self.email, self.password]):
            raise ValueError("请在.env文件中设置EMAIL和EMAIL_PASSWORD")

class EmailSender:
    """邮件发送类"""
    
    def __init__(self, config: EmailConfig):
        """
        初始化邮件发送器
        
        Args:
            config: EmailConfig实例，包含邮件服务器配置
        """
        self.config = config
        self.smtp_client = None
        self.max_retries = 3
        self.retry_delay = 2  # 重试延迟（秒）
        
    def __enter__(self):
        """上下文管理器入口"""
        self.connect()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.close()
        
    def connect(self) -> bool:
        """连接到SMTP服务器"""
        for attempt in range(self.max_retries):
            try:
                self.smtp_client = smtplib.SMTP_SSL(
                    self.config.smtp_server, 
                    self.config.smtp_port
                )
                self.smtp_client.login(self.config.email, self.config.password)
                logger.info("成功连接到SMTP服务器")
                return True
            except Exception as e:
                logger.error(f"第{attempt + 1}次连接失败: {str(e)}")
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)
                    continue
                return False

    def send_email(self, 
                  receiver: Union[str, List[str]],
                  subject: str,
                  content: str,
                  attachments: Optional[List[str]] = None,
                  html_content: Optional[str] = None,
                  sender_name: str = "",
                  receiver_name: str = "",
                  cc: Optional[List[str]] = None,
                  bcc: Optional[List[str]] = None) -> bool:
        """
        发送邮件
        
        Args:
            receiver: 收件人邮箱或邮箱列表
            subject: 邮件主题
            content: 邮件文本内容
            attachments: 附件路径列表
            html_content: HTML格式的邮件内容
            sender_name: 发件人名称
            receiver_name: 收件人名称
            cc: 抄送列表
            bcc: 密送列表
            
        Returns:
            bool: 发送是否成功
        """
        try:
            # 创建邮件对象
            message = MIMEMultipart('alternative')
            message["Subject"] = Header(subject)
            message["From"] = Header(f"{sender_name}<{self.config.email}>" if sender_name else self.config.email)
            
            # 处理收件人列表
            receivers = [receiver] if isinstance(receiver, str) else receiver
            message["To"] = Header(", ".join(receivers))
            
            # 处理抄送和密送
            if cc:
                message["Cc"] = Header(", ".join(cc))
                receivers.extend(cc)
            if bcc:
                receivers.extend(bcc)
            
            # 添加文本内容
            message.attach(MIMEText(content, "plain", "utf-8"))
            
            # 如果有HTML内容，添加HTML部分
            if html_content:
                message.attach(MIMEText(html_content, "html", "utf-8"))

            # 添加附件
            if attachments:
                for attachment_path in attachments:
                    self._add_attachment(message, attachment_path)

            # 发送邮件（带重试机制）
            for attempt in range(self.max_retries):
                try:
                    self.smtp_client.sendmail(
                        self.config.email, 
                        receivers, 
                        message.as_string()
                    )
                    logger.info(f"邮件已成功发送给 {', '.join(receivers)}")
                    return True
                except Exception as e:
                    logger.error(f"第{attempt + 1}次发送失败: {str(e)}")
                    if attempt < self.max_retries - 1:
                        time.sleep(self.retry_delay)
                        continue
                    return False
            
        except Exception as e:
            logger.error(f"发送邮件时发生错误: {str(e)}")
            return False
            
    def _add_attachment(self, message: MIMEMultipart, file_path: str):
        """添加附件到邮件"""
        try:
            path = Path(file_path)
            if not path.exists():
                logger.warning(f"附件不存在: {file_path}")
                return

            with open(path, "rb") as f:
                file_content = f.read()
                
            # 根据文件类型创建不同的MIME对象
            if path.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif']:
                attachment = MIMEImage(file_content)
            else:
                attachment = MIMEApplication(file_content)
                
            attachment.add_header(
                "Content-Disposition", 
                "attachment", 
                filename=path.name
            )
            message.attach(attachment)
            logger.debug(f"成功添加附件: {path.name}")
            
        except Exception as e:
            logger.error(f"添加附件时发生错误: {str(e)}")
            
    def close(self):
        """关闭SMTP连接"""
        if self.smtp_client:
            try:
                self.smtp_client.quit()
                logger.info("SMTP连接已关闭")
            except Exception as e:
                logger.error(f"关闭SMTP连接时发生错误: {str(e)}")

def main():
    """主函数"""
    try:
        # 加载配置
        config = EmailConfig()
        
        # 使用上下文管理器创建EmailSender实例
        with EmailSender(config) as sender:
            # 发送邮件
            success = sender.send_email(
                receiver=["dong.cheng@ythyjx.com"],
                subject="结业证书",
                content="Tom，这是夜曲编程的结业证书，望查收~",
                attachments=[r"D:\Python\Practices\自动发邮件\certificate_tom.png"],
                html_content="<h1>恭喜完成课程！</h1><p>这是您的结业证书，请查收。</p>",
                sender_name="夜曲编程",
                receiver_name="Tom",
                cc=["support@ythyjx.com"]
            )
            
            if success:
                logger.info("邮件发送成功！")
            else:
                logger.error("邮件发送失败！")
                
    except Exception as e:
        logger.error(f"程序执行出错: {str(e)}")

if __name__ == "__main__":
    main()