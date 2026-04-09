import os
import re
import ssl
import smtplib
import traceback
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import streamlit as st


# =========================
# 基础配置
# =========================
st.set_page_config(
    page_title="供应商信息采集 / Supplier Information Upload",
    page_icon="📸",
    layout="centered",
)

MAX_FILE_SIZE_MB = 5
MAX_TOTAL_FILES = 5
MAX_DETAIL_FILES = 3
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png"}


# =========================
# Session State 初始化
# =========================
if "uploader_nonce" not in st.session_state:
    st.session_state["uploader_nonce"] = 0

if "pending_submission" not in st.session_state:
    st.session_state["pending_submission"] = None

if "flash_success" not in st.session_state:
    st.session_state["flash_success"] = ""


# =========================
# 样式：更适合手机端
# =========================
st.markdown(
    """
    <style>
    .main {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }
    h1, h2, h3 {
        line-height: 1.2;
    }
    div[data-testid="stFormSubmitButton"] button {
        width: 100%;
        min-height: 3rem;
        font-size: 1.05rem;
        font-weight: 700;
        border-radius: 12px;
    }
    div[data-testid="stButton"] button {
        width: 100%;
        min-height: 3rem;
        font-size: 1.0rem;
        font-weight: 700;
        border-radius: 12px;
    }
    .small-note {
        color: #666;
        font-size: 0.92rem;
        margin-top: -0.2rem;
        margin-bottom: 0.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# 工具函数
# =========================
def normalize_recipients(value):
    """支持 list / tuple / 字符串。"""
    if value is None:
        return []

    if isinstance(value, (list, tuple)):
        return [str(x).strip() for x in value if str(x).strip()]

    if isinstance(value, str):
        parts = re.split(r"[,\n;]+", value)
        return [p.strip() for p in parts if p.strip()]

    return []


def load_email_config():
    """
    从 Streamlit Secrets 读取邮箱配置。
    需要在 .streamlit/secrets.toml 或 Streamlit Cloud Secrets 中填写。
    """
    try:
        email_cfg = st.secrets["email"]
    except Exception:
        raise ValueError(
            "未找到邮箱配置。请先在 .streamlit/secrets.toml "
            "或 Streamlit Cloud 的 Secrets 中填写 [email] 配置。"
        )

    config = {
        "smtp_host": str(email_cfg.get("smtp_host", "")).strip(),
        "smtp_port": int(email_cfg.get("smtp_port", 465)),
        "use_ssl": bool(email_cfg.get("use_ssl", True)),
        "sender_email": str(email_cfg.get("sender_email", "")).strip(),
        "sender_password": str(email_cfg.get("sender_password", "")).strip(),
        "sender_name": str(email_cfg.get("sender_name", "Supplier Upload Bot")).strip(),
        "recipients": normalize_recipients(email_cfg.get("recipients", [])),
    }

    missing = []
    for key in ["smtp_host", "smtp_port", "sender_email", "sender_password"]:
        if not config.get(key):
            missing.append(key)

    if not config["recipients"]:
        missing.append("recipients")

    if missing:
        raise ValueError("邮箱配置不完整，请补齐： " + ", ".join(missing))

    return config


def sanitize_filename(filename: str) -> str:
    """去掉非法字符，保留原扩展名。"""
    base = os.path.basename(filename or "image.jpg")
    base = base.strip().replace(" ", "_")
    base = re.sub(r'[\\/:*?"<>|\r\n]+', "_", base)
    return base or "image.jpg"


def get_extension(filename: str) -> str:
    return os.path.splitext(filename)[1].lower()


def ext_to_subtype(ext: str):
    if ext in [".jpg", ".jpeg"]:
        return "jpeg"
    if ext == ".png":
        return "png"
    return None


def validate_and_build_attachment(uploaded_file, prefix: str):
    """
    全程只在内存中处理：
    读取 UploadedFile 的字节数据 -> 检查类型和大小 -> 生成重命名后的附件信息
    """
    if uploaded_file is None:
        raise ValueError("存在未上传的文件。")

    original_name = sanitize_filename(uploaded_file.name)
    ext = get_extension(original_name)

    if ext not in ALLOWED_EXTENSIONS:
        raise ValueError(
            f"不支持的文件类型：{original_name}。仅支持 JPG / JPEG / PNG。"
        )

    file_bytes = uploaded_file.getvalue()
    size_mb = len(file_bytes) / (1024 * 1024)

    if size_mb > MAX_FILE_SIZE_MB:
        raise ValueError(
            f"文件过大：{original_name}（{size_mb:.2f} MB），"
            f"请控制在 {MAX_FILE_SIZE_MB} MB 以内。"
        )

    subtype = ext_to_subtype(ext)
    if not subtype:
        raise ValueError(f"无法识别图片类型：{original_name}")

    new_name = f"{prefix}_{original_name}"

    return {
        "filename": new_name,
        "bytes": file_bytes,
        "maintype": "image",
        "subtype": subtype,
        "size_mb": round(size_mb, 2),
    }


def collect_attachments(card_file, product_file, detail_files):
    """
    命名规则：
    01_名片_原文件名
    02_产品_原文件名
    03_其他_原文件名
    04_其他_原文件名
    05_其他_原文件名
    """
    if card_file is None:
        raise ValueError("请上传名片照片 / Please upload a business card photo.")

    if product_file is None:
        raise ValueError("请上传产品照片 / Please upload a product photo.")

    detail_files = detail_files or []

    if len(detail_files) > MAX_DETAIL_FILES:
        raise ValueError(
            f"其他细节图片最多上传 {MAX_DETAIL_FILES} 张 / "
            f"Up to {MAX_DETAIL_FILES} detail photos."
        )

    attachments = []
    attachments.append(validate_and_build_attachment(card_file, "01_名片"))
    attachments.append(validate_and_build_attachment(product_file, "02_产品"))

    for idx, detail in enumerate(detail_files, start=3):
        attachments.append(validate_and_build_attachment(detail, f"{idx:02d}_其他"))

    if len(attachments) > MAX_TOTAL_FILES:
        raise ValueError(
            f"最多允许上传 {MAX_TOTAL_FILES} 张图片 / Up to {MAX_TOTAL_FILES} images."
        )

    return attachments


def build_subject(company_short_name: str, upload_time_str: str) -> str:
    company_short_name = re.sub(r"[\r\n]+", " ", (company_short_name or "").strip())
    company_short_name = company_short_name if company_short_name else "Unknown Supplier"
    return f"[展会供应商采集 / Supplier Upload] {company_short_name} | {upload_time_str}"


def build_body(company_short_name: str, remarks: str, upload_time_str: str, attachments: list) -> str:
    company_display = company_short_name.strip() if company_short_name.strip() else "未填写 / Not provided"
    remarks_display = remarks.strip() if remarks.strip() else "未填写 / Not provided"

    lines = [
        "您好，",
        "",
        "这是来自展会现场的供应商图片采集邮件。",
        "This is a supplier image submission collected at the exhibition.",
        "",
        f"上传时间 / Upload time: {upload_time_str}",
        f"公司简称 / Company short name: {company_display}",
        f"备注 / Notes: {remarks_display}",
        f"附件数量 / Number of attachments: {len(attachments)}",
        "",
        "附件清单 / Attachment list:",
    ]

    for item in attachments:
        lines.append(f"- {item['filename']}")

    lines.extend(
        [
            "",
            "感谢您分享信息，我们会尽快查看并跟进。",
            "Thank you for sharing your information. We will review it and follow up as soon as possible.",
        ]
    )

    return "\n".join(lines)


def mask_password(pwd: str) -> str:
    if not pwd:
        return "(empty)"
    if len(pwd) <= 4:
        return "*" * len(pwd)
    return pwd[:2] + "*" * (len(pwd) - 4) + pwd[-2:]


def send_email(subject: str, body: str, attachments: list):
    """
    发送邮件。
    ✅ 增强版：支持状态追踪、DSN 回执、本地日志记录、智能提示
    """
    cfg = load_email_config()

    # 初始化本地日志路径
    LOG_FILE = Path("logs/email_send.log")
    LOG_FILE.parent.mkdir(exist_ok=True)

    # 生成唯一 Message-ID（便于后续追踪）
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")[:-3]
    message_id = f"<{timestamp}.{id(subject)}@{cfg['smtp_host']}>"

    msg = EmailMessage()
    sender_name = cfg["sender_name"]
    sender_email = cfg["sender_email"]
    recipients = cfg["recipients"]

    msg["From"] = formataddr((sender_name, sender_email)) if sender_name else sender_email
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg["Message-ID"] = message_id  # ✅ 唯一追踪ID

    # ✅ 添加 DSN 回执请求头（部分服务器支持）
    msg["Disposition-Notification-To"] = sender_email
    msg["Return-Receipt-To"] = sender_email
    msg["X-Priority"] = "3"  # 普通优先级

    msg.set_content(body)

    for item in attachments:
        msg.add_attachment(
            item["bytes"],
            maintype=item["maintype"],
            subtype=item["subtype"],
            filename=item['filename'],
        )

    # 🎯 智能风险提示预计算
    warnings = []
    # 检查企业邮箱（非个人域名）
    personal_domains = {"qq.com", "163.com", "126.com", "gmail.com", "outlook.com", "hotmail.com", "sina.com",
                        "139.com"}
    enterprise_recipients = [r for r in recipients if "@" in r and r.split("@")[1].lower() not in personal_domains]
    if enterprise_recipients:
        warnings.append(
            f"⚠️ 检测到企业邮箱 ({', '.join(enterprise_recipients)})，可能有审计延迟，建议 10 分钟后未收到再排查")

    total_attach_size = sum(item.get("size_mb", 0) for item in attachments)
    if total_attach_size > 5:
        warnings.append(f"⚠️ 附件总计 {total_attach_size:.1f} MB，可能被收件方反垃圾系统延迟扫描")

    send_result = {}  # 用于记录 send_message 返回结果

    try:
        if cfg["use_ssl"]:
            with smtplib.SMTP_SSL(
                    cfg["smtp_host"],
                    cfg["smtp_port"],
                    context=ssl.create_default_context(),
                    timeout=30
            ) as server:
                server.login(sender_email, cfg["sender_password"])
                # ✅ 解析 send_message 返回的详细状态码
                # 返回值: {失败邮箱: (错误码, 错误信息)}，空字典表示全部成功
                send_result = server.send_message(msg, to_addrs=recipients)

        else:
            with smtplib.SMTP(cfg["smtp_host"], cfg["smtp_port"], timeout=30) as server:
                server.ehlo()
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
                server.login(sender_email, cfg["sender_password"])
                # ✅ 解析 send_message 返回的详细状态码
                send_result = server.send_message(msg, to_addrs=recipients)

    finally:
        # ✅ 记录到本地日志文件（无论成功失败都记录）
        try:
            log_entry = {
                "timestamp": datetime.now().isoformat(),
                "message_id": message_id,
                "sender": sender_email,
                "recipients": recipients,
                "subject": subject,
                "attachments_count": len(attachments),
                "success": len(send_result) == 0,
                "failed_recipients": {k: str(v) for k, v in send_result.items()},
                "warnings": warnings,
            }
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"{log_entry}\n")
        except Exception:
            pass  # 静默处理日志写入失败


def clear_pending_submission():
    st.session_state["pending_submission"] = None


# =========================
# 页面顶部
# =========================
if st.session_state["flash_success"]:
    st.success(st.session_state["flash_success"])
    st.session_state["flash_success"] = ""

st.title("📸 供应商信息采集 / Supplier Information Upload")

st.markdown(
    """
    <div class="small-note">
    建议填写公司名简称，方便我们更快识别并优先跟进您。<br>
    We recommend entering your company short name so we can identify and follow up with you faster.
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# 主表单
# =========================
nonce = st.session_state["uploader_nonce"]

with st.form(key=f"upload_form_{nonce}", clear_on_submit=False):
    company_short_name = st.text_input(
        "公司名简称（选填） / Company short name (optional)",
        placeholder="例如 / e.g. ABC, XX Metal, Sunshine",
        help="建议填写，方便我们更快识别并优先跟进您 / Recommended for faster identification and follow-up.",
    )

    card_file = st.file_uploader(
        "1) 名片 / Business Card（1 张）",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=False,
        help="请上传 1 张名片照片 / Please upload 1 business card photo.",
    )

    product_file = st.file_uploader(
        "2) 产品图 / Product Photo（1 张）",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=False,
        help="请上传 1 张产品照片 / Please upload 1 product photo.",
    )

    detail_files = st.file_uploader(
        "3) 其他细节 / Other Details（0–3 张）",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        help="可上传其他细节图片，最多 3 张 / You may upload up to 3 detail photos.",
    )

    remarks = st.text_area(
        "备注（选填） / Notes (optional)",
        placeholder=(
            "补充信息，例如贵司产品类别、可加工材质、关键生产工艺、表面处理工艺等。\n"
            "Add anynotes: product categories, materials, capability, surface treatments etc."
        ),
        height=130,
    )

    review_clicked = st.form_submit_button(
        "检查并预览 / Review before sending",
        use_container_width=True,
    )

    if review_clicked:
        try:
            attachments = collect_attachments(card_file, product_file, detail_files)
            upload_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            st.session_state["pending_submission"] = {
                "company_short_name": company_short_name.strip(),
                "remarks": remarks.strip(),
                "attachments": attachments,
                "upload_time_str": upload_time_str,
            }

        except Exception as e:
            clear_pending_submission()
            st.error(f"校验失败 / Validation failed: {e}")

# =========================
# 发送前预览
# =========================
pending = st.session_state.get("pending_submission")

if pending:
    st.subheader("📋 发送前预览 / Review before sending")

    st.write("**邮件主题 / Email subject**")
    st.code(build_subject(pending["company_short_name"], pending["upload_time_str"]), language=None)

    company_preview = pending["company_short_name"] if pending["company_short_name"] else "未填写 / Not provided"
    remarks_preview = pending["remarks"] if pending["remarks"] else "未填写 / Not provided"

    st.write(f"**公司简称 / Company short name**：{company_preview}")
    st.write(f"**备注 / Notes**：{remarks_preview}")
    st.write(f"**上传时间 / Upload time**：{pending['upload_time_str']}")
    st.write(f"**附件数量 / Attachments**：{len(pending['attachments'])} 张")

    st.info(
        "确认无误后再发送。\nPlease confirm everything is correct before sending."
    )

    col1, col2 = st.columns(2)

    with col1:
        confirm_send = st.button(
            "确认发送 / Confirm and Send",
            type="primary",
            use_container_width=True,
        )

    with col2:
        cancel_send = st.button(
            "取消发送 / Cancel",
            use_container_width=True,
        )

    if cancel_send:
        clear_pending_submission()
        st.rerun()

    if confirm_send:
        try:
            subject = build_subject(
                pending["company_short_name"],
                pending["upload_time_str"]
            )
            body = build_body(
                pending["company_short_name"],
                pending["remarks"],
                pending["upload_time_str"],
                pending["attachments"],
            )

            send_email(subject, body, pending["attachments"])

            clear_pending_submission()
            st.session_state["flash_success"] = (
                "✅ 上传成功，邮件已发送。谢谢！ / Upload successful, email sent. Thank you!"
            )

            # 刷新上传控件，避免上一家供应商内容残留
            st.session_state["uploader_nonce"] += 1
            st.rerun()

        except smtplib.SMTPAuthenticationError:
            st.error("❌ 认证失败！请检查：1. QQ邮箱账号 2. 授权码（不是登录密码） 3. SMTP 是否已开启")

        except smtplib.SMTPRecipientsRefused as e:
            st.error(f"❌ 收件人被拒绝：{e}")

        except smtplib.SMTPSenderRefused as e:
            st.error(f"❌ 发件人被拒绝：{e}")

        except smtplib.SMTPConnectError as e:
            st.error(f"❌ SMTP 连接失败：{e}")

        except smtplib.SMTPServerDisconnected as e:
            st.error(f"❌ SMTP 服务器断开连接：{e}")

        except smtplib.SMTPException as e:
            st.error(f"❌ SMTP 发送失败：{type(e).__name__}: {e}")

        except Exception as e:
            st.error(f"❌ 发送失败：{type(e).__name__}: {e}")

else:
    st.caption(
        '填写信息后，先点击"检查并预览"，确认无误后再发送。'
        'After filling in the form, click "Review before sending" first, then confirm and send.'
    )


st.divider()
st.caption(
    "本工具全程以内存处理图片，不写入本地磁盘。"
    "This tool processes images in memory only and does not save them to local disk."
)
