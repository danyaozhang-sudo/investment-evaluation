# 邮件投递指南

投资评估报告完成后，用户可能要求通过邮件发送。

## 首选方式：Python smtplib（可靠）

当 `himalaya` CLI 的 IMAP保存到已发送 环节失败时，退路方案：

```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# 配置
smtp_host = 'smtp.exmail.qq.com'
smtp_port = 465
username = 'dazongguan@janoenergy.cn'
password = '<密码>'

msg = MIMEMultipart('mixed')
msg['From'] = '大总关 <dazongguan@janoenergy.cn>'
msg['To'] = 'dazongguan@janoenergy.cn'
msg['Subject'] = '<主题>'

# 正文
text = MIMEText('<正文>', 'plain', 'utf-8')
msg.attach(text)

# 附件
with open('<附件路径>', 'rb') as f:
    attachment = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
    attachment.set_payload(f.read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename='<显示文件名>')
msg.attach(attachment)

# 发送
with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
    server.login(username, password)
    server.sendmail(from_addr, [to_addr], msg.as_string())
```

## 备选方式：himalaya CLI

```bash
cat << 'MAIL' | himalaya template send
From: sender@example.com
To: recipient@example.com
Subject: 报告

正文...

<#part filename=/path/to/report.docx><#/part>
MAIL
```

**注意**：himalaya 发送成功后尝试保存到已发送文件夹，若IMAP连接不稳定会导致报错（`cannot add IMAP message`），但SMTP投递通常已成功。此时用Python smtplib更可靠。

## 收件人地址

用户邮箱：`dazongguan@janoenergy.cn`

## 邮件标题格式

`广东潮州XXX项目投资评估报告`
