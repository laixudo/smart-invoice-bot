"""国内常见企业邮 SMTP 预设 + 发信。"""

from __future__ import annotations

import mimetypes
import smtplib
import ssl
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable

# 仅保留 QQ / 腾讯企业邮
SMTP_PRESETS: dict[str, dict] = {
    "qq": {
        "label": "QQ邮箱",
        "host": "smtp.qq.com",
        "port": 465,
        "use_ssl": True,
    },
    "tencent_exmail": {
        "label": "腾讯企业邮",
        "host": "smtp.exmail.qq.com",
        "port": 465,
        "use_ssl": True,
    },
}


@dataclass
class MailPayload:
    subject: str
    body: str
    to: list[str]
    cc: list[str]
    attachments: list[Path]
    from_addr: str


@dataclass
class SmtpConfig:
    provider: str
    host: str
    port: int
    use_ssl: bool
    username: str
    auth_code: str  # 客户端授权码


def resolve_smtp(email_cfg: dict, local_cfg: dict) -> SmtpConfig:
    provider = (local_cfg.get("provider") or email_cfg.get("provider") or "tencent_exmail").strip()
    preset = SMTP_PRESETS.get(provider, SMTP_PRESETS["qq"])
    host = local_cfg.get("host") or email_cfg.get("host") or preset["host"]
    port = int(local_cfg.get("port") or email_cfg.get("port") or preset["port"] or 465)
    use_ssl = local_cfg.get("use_ssl")
    if use_ssl is None:
        use_ssl = email_cfg.get("use_ssl")
    if use_ssl is None:
        use_ssl = bool(preset.get("use_ssl", True))
    username = (local_cfg.get("username") or email_cfg.get("username") or "").strip()
    auth_code = (local_cfg.get("auth_code") or email_cfg.get("auth_code") or "").strip()
    if not host:
        raise ValueError("未配置 SMTP 服务器（host）")
    if not username:
        raise ValueError("未配置发件邮箱（username）")
    if not auth_code:
        raise ValueError("未配置客户端授权码（auth_code），请写入 config.local.yaml")
    return SmtpConfig(
        provider=provider,
        host=host,
        port=port,
        use_ssl=bool(use_ssl),
        username=username,
        auth_code=auth_code,
    )


def send_smtp(smtp: SmtpConfig, mail: MailPayload) -> None:
    msg = EmailMessage()
    msg["Subject"] = mail.subject
    msg["From"] = mail.from_addr or smtp.username
    msg["To"] = ", ".join(mail.to)
    if mail.cc:
        msg["Cc"] = ", ".join(mail.cc)
    msg.set_content(mail.body)

    for path in mail.attachments:
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"附件不存在: {path}")
        ctype, encoding = mimetypes.guess_type(str(path))
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    recipients = list(mail.to) + list(mail.cc)
    context = ssl.create_default_context()
    if smtp.use_ssl:
        with smtplib.SMTP_SSL(smtp.host, smtp.port, context=context, timeout=60) as server:
            server.login(smtp.username, smtp.auth_code)
            server.send_message(msg, from_addr=smtp.username, to_addrs=recipients)
    else:
        with smtplib.SMTP(smtp.host, smtp.port, timeout=60) as server:
            server.ehlo()
            server.starttls(context=context)
            server.login(smtp.username, smtp.auth_code)
            server.send_message(msg, from_addr=smtp.username, to_addrs=recipients)


def total_attachment_mb(paths: Iterable[Path]) -> float:
    return sum(Path(p).stat().st_size for p in paths) / (1024 * 1024)
