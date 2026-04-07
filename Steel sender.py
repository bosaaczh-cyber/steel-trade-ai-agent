# -*- coding: utf-8 -*-
"""
Abter Steel — 阿里云邮件推送发送模块
═══════════════════════════════════════
对接 steel_master.py 生成的开发信，直接调用阿里云 DirectMail API 发送。

使用方式：
  python steel_sender.py --mode test    # 先发1封测试到自己邮箱
  python steel_sender.py --mode send    # 正式发送（读取 final_emails.json）
  python steel_sender.py --mode status  # 查询发送统计

.env 需要新增：
  ALIYUN_ACCESS_KEY_ID=your_key_here
  ALIYUN_ACCESS_KEY_SECRET=你的Secret
  ALIYUN_FROM_ADDRESS=bosaa@abtersteeltube.com
  ALIYUN_FROM_ALIAS=Bosaa
  TEST_EMAIL=你自己的邮箱（用来收测试邮件）
"""

import os
import json
import csv
import time
import hmac
import base64
import hashlib
import uuid
import argparse
from datetime import datetime, date
from pathlib import Path
from urllib import parse, request as urlrequest
from urllib.error import URLError
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════
# 配置（从 .env 读取）
# ══════════════════════════════════════════
ACCESS_KEY_ID = os.environ.get("ALIYUN_ACCESS_KEY_ID", "")
ACCESS_KEY_SECRET = os.environ.get("ALIYUN_ACCESS_KEY_SECRET", "")
FROM_ADDRESS = os.environ.get("ALIYUN_FROM_ADDRESS", "")
FROM_ALIAS        = os.environ.get("ALIYUN_FROM_ALIAS", "Bosaa")
TEST_EMAIL        = os.environ.get("TEST_EMAIL", "")

# 阿里云 DirectMail API 端点（华东1 杭州）
API_ENDPOINT = "https://dm.aliyuncs.com"

# 发送间隔（秒）— 避免触发频率限制，建议不低于 0.5
SEND_INTERVAL = 0.8

# ══════════════════════════════════════════
# 阿里云 API 签名（不依赖 SDK，纯标准库）
# ══════════════════════════════════════════
def _percent_encode(s: str) -> str:
    return parse.quote(str(s), safe="")

def _sign(params: dict, secret: str) -> str:
    """生成阿里云 API 签名"""
    sorted_params = sorted(params.items())
    query_string  = "&".join([f"{_percent_encode(k)}={_percent_encode(v)}"
                               for k, v in sorted_params])
    string_to_sign = f"GET&{_percent_encode('/')}&{_percent_encode(query_string)}"
    key = (secret + "&").encode("utf-8")
    hashed = hmac.new(key, string_to_sign.encode("utf-8"), hashlib.sha1)
    return base64.b64encode(hashed.digest()).decode("utf-8")

def _call_api(action: str, extra_params: dict) -> dict:
    """调用阿里云 DirectMail API"""
    params = {
        "Format":           "JSON",
        "Version":          "2015-11-23",
        "AccessKeyId":      ACCESS_KEY_ID,
        "SignatureMethod":  "HMAC-SHA1",
        "Timestamp":        datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "SignatureVersion": "1.0",
        "SignatureNonce":   str(uuid.uuid4()),
        "Action":           action,
        **extra_params,
    }
    params["Signature"] = _sign(params, ACCESS_KEY_SECRET)
    url = f"{API_ENDPOINT}/?{parse.urlencode(params)}"
    try:
        with urlrequest.urlopen(url, timeout=15) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except URLError as e:
        return {"Error": str(e)}

# ══════════════════════════════════════════
# 核心发送函数
# ══════════════════════════════════════════
def send_single_email(to_address: str, subject: str, body: str,
                      to_alias: str = "") -> dict:
    """
    发送单封邮件
    返回：{"success": True/False, "message_id": "...", "error": "..."}
    """
    if not ACCESS_KEY_SECRET:
        return {"success": False, "error": "ALIYUN_ACCESS_KEY_SECRET 未配置"}
    if not to_address or "@" not in to_address:
        return {"success": False, "error": f"无效邮箱地址：{to_address}"}

    params = {
        "AccountName":   FROM_ADDRESS,
        "FromAlias":     FROM_ALIAS,
        "AddressType":   "1",          # 1=随机账号（外部），0=触发邮件
        "TagName":       "abter_steel_outreach",
        "ReplyToAddress": "false",
        "ToAddress":     to_address,
        "Subject":       subject,
        "TextBody":      body,         # 纯文本，兼容性最好
    }
    if to_alias:
        params["ToAlias"] = to_alias

    result = _call_api("SingleSendMail", params)

    if "EnvId" in result or "RequestId" in result:
        return {
            "success":    True,
            "request_id": result.get("RequestId", ""),
            "env_id":     result.get("EnvId", ""),
        }
    else:
        error_msg = result.get("Message", str(result))
        error_code = result.get("Code", "UNKNOWN")
        return {
            "success": False,
            "error":   f"[{error_code}] {error_msg}",
            "raw":     result,
        }

# ══════════════════════════════════════════
# 从开发信文本解析 Subject + Body
# ══════════════════════════════════════════
def parse_email_text(text: str) -> tuple:
    """
    输入：steel_master.py 生成的完整邮件文本
    输出：(subject, body)
    """
    subject    = ""
    body_lines = []
    for line in text.split("\n"):
        line = line.strip().replace("**", "")
        if line.lower().startswith("subject:"):
            subject = line.split(":", 1)[-1].strip()
        elif line:
            body_lines.append(line)
    body = "\n".join(body_lines)
    return subject, body

# ══════════════════════════════════════════
# 发送日志
# ══════════════════════════════════════════
LOG_FILE = "send_log.json"

def load_log() -> list:
    if not Path(LOG_FILE).exists():
        return []
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_log(log: list):
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

def append_log(entry: dict):
    log = load_log()
    log.append(entry)
    save_log(log)

# ══════════════════════════════════════════
# 模式 1：test — 发1封测试邮件到自己
# ══════════════════════════════════════════
def mode_test():
    print("\n🧪 测试模式：发1封邮件到你自己\n")

    if not TEST_EMAIL:
        test_addr = input("  输入你的收件邮箱：").strip()
    else:
        test_addr = TEST_EMAIL
        print(f"  收件地址：{test_addr}")

    subject = "【测试】Abter Steel 邮件系统连通性测试"
    body    = (
        f"Hi,\n\n"
        f"This is a test email from Abter Steel outreach system.\n\n"
        f"If you received this, the system is working correctly.\n\n"
        f"Send time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"From: {FROM_ADDRESS}\n\n"
        f"Best regards,\n"
        f"Abter Steel\n"
        f"sales@abter-steel.com"
    )

    print(f"\n  发送中...")
    result = send_single_email(test_addr, subject, body)

    if result["success"]:
        print(f"  ✅ 发送成功！RequestID: {result.get('request_id','')}")
        print(f"  📬 请检查 {test_addr} 的收件箱（注意查垃圾邮件）")
    else:
        print(f"  ❌ 发送失败：{result['error']}")
        print(f"\n  常见原因：")
        print(f"  · ALIYUN_ACCESS_KEY_SECRET 未填写或填错")
        print(f"  · 发信地址 {FROM_ADDRESS} 未在阿里云控制台验证")
        print(f"  · AccessKey 没有 DirectMail 权限")

# ══════════════════════════════════════════
# 模式 2：send — 从 CSV 批量发送
# ══════════════════════════════════════════
def mode_send():
    """
    读取 steel_master.py 生成的 CSV 批量发送。
    CSV 格式（steel_master.py 导出的）：
      收件人邮箱 | 联系人姓名 | 公司 | 邮件主题 | 邮件正文
    """
    print("\n📤 批量发送模式\n")

    # 找可用的 CSV 文件
    csv_files = sorted(Path(".").glob("abter_steel_full_*.csv")) + \
                sorted(Path(".").glob("followup_send_*.csv"))

    if not csv_files:
        # 让用户手动输入路径
        csv_path = input("  找不到CSV文件，请输入文件路径：").strip()
        csv_files = [Path(csv_path)]
    else:
        print("  找到以下可发送文件：")
        for i, f in enumerate(csv_files, 1):
            print(f"  {i}. {f.name}")
        idx = int(input("\n  选择哪个文件（序号）：").strip()) - 1
        csv_files = [csv_files[idx]]

    csv_path = csv_files[0]
    print(f"\n  读取文件：{csv_path}")

    # 读取 CSV
    recipients = []
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 兼容两种列名格式
            email   = row.get("收件人邮箱") or row.get("email", "")
            contact = row.get("联系人姓名") or row.get("contact", "")
            company = row.get("公司")        or row.get("company", "")
            subject = row.get("邮件主题")    or row.get("subject", "")
            body    = row.get("邮件正文")    or row.get("body", "")
            if email:
                recipients.append({
                    "email": email.strip(),
                    "contact": contact.strip(),
                    "company": company.strip(),
                    "subject": subject.strip(),
                    "body":    body.strip(),
                })

    if not recipients:
        print("  ❌ CSV 为空或格式不对")
        return

    # 过滤无效邮箱
    valid   = [r for r in recipients if "@" in r["email"]]
    invalid = [r for r in recipients if "@" not in r["email"]]

    print(f"\n  总计：{len(recipients)} 条")
    print(f"  有效邮箱：{len(valid)} 条")
    if invalid:
        print(f"  ⚠️  无效邮箱（跳过）：{len(invalid)} 条")
        for r in invalid:
            print(f"     · {r['company']} — {r['email']}")

    if not valid:
        print("\n  ❌ 没有有效收件人，退出")
        return

    # 预估时间
    est_seconds = len(valid) * SEND_INTERVAL
    est_min     = int(est_seconds / 60)
    print(f"\n  预计发送时间：约 {est_min} 分钟（间隔 {SEND_INTERVAL}s/封）")
    print(f"  日额度消耗：{len(valid)}/8000")

    print(f"\n  前3条预览：")
    for r in valid[:3]:
        print(f"  · {r['company']:<25} → {r['email']}")
        print(f"    Subject: {r['subject'][:50]}")

    confirm = input(f"\n  确认发送 {len(valid)} 封邮件？(y/n) ").strip().lower()
    if confirm != "y":
        print("  已取消")
        return

    # 开始发送
    print(f"\n  {'─'*55}")
    success_count = 0
    fail_count    = 0
    today_str     = str(date.today())

    for i, r in enumerate(valid, 1):
        print(f"  [{i:03d}/{len(valid)}] {r['company']:<25} → {r['email'][:30]}...", end=" ")

        result = send_single_email(
            to_address = r["email"],
            subject    = r["subject"],
            body       = r["body"],
            to_alias   = r["contact"],
        )

        if result["success"]:
            print("✅")
            success_count += 1
            status = "success"
        else:
            print(f"❌ {result.get('error','')[:40]}")
            fail_count += 1
            status = "failed"

        # 记录日志
        append_log({
            "date":       today_str,
            "time":       datetime.now().strftime("%H:%M:%S"),
            "company":    r["company"],
            "contact":    r["contact"],
            "email":      r["email"],
            "subject":    r["subject"],
            "status":     status,
            "error":      result.get("error", ""),
            "request_id": result.get("request_id", ""),
        })

        # 间隔（避免限流）
        if i < len(valid):
            time.sleep(SEND_INTERVAL)

    print(f"\n  {'═'*55}")
    print(f"  ✅ 发送成功：{success_count} 封")
    print(f"  ❌ 发送失败：{fail_count} 封")
    print(f"  📋 发送日志已记录 → {LOG_FILE}")
    print(f"  {'═'*55}\n")

    # 自动写入跟进数据库
    if success_count > 0:
        _sync_to_followup_db(valid, today_str)

# ══════════════════════════════════════════
# 模式 3：status — 查看发送统计
# ══════════════════════════════════════════
def mode_status():
    print("\n📊 发送统计\n")
    log = load_log()
    if not log:
        print("  ⚠️  还没有发送记录")
        return

    total   = len(log)
    success = sum(1 for l in log if l["status"] == "success")
    failed  = sum(1 for l in log if l["status"] == "failed")

    # 按日期分组
    by_date = {}
    for entry in log:
        d = entry.get("date", "unknown")
        by_date.setdefault(d, {"success": 0, "failed": 0})
        by_date[d][entry["status"]] = by_date[d].get(entry["status"], 0) + 1

    print(f"  总发送量：{total} 封")
    print(f"  成功：{success}  |  失败：{failed}")
    print(f"  成功率：{success/total*100:.1f}%\n" if total else "")

    print(f"  按日期：")
    for d in sorted(by_date.keys(), reverse=True):
        s = by_date[d]["success"]
        f = by_date[d].get("failed", 0)
        print(f"  {d}  ✅{s} ❌{f}")

    # 失败明细
    failures = [l for l in log if l["status"] == "failed"]
    if failures:
        print(f"\n  失败明细（最近5条）：")
        for l in failures[-5:]:
            print(f"  · {l['company']:<20} {l['email']:<30} {l.get('error','')[:40]}")

# ══════════════════════════════════════════
# 发送成功后自动写入跟进数据库
# ══════════════════════════════════════════
def _sync_to_followup_db(sent_list: list, sent_date: str):
    """把成功发送的邮件自动录入跟进数据库"""
    try:
        from steel_followup import load_db, save_db, add_lead
        db = load_db()
        # 获取已有邮箱，避免重复录入
        existing_emails = {l.get("email", "") for l in db["leads"]}
        added = 0
        for r in sent_list:
            if r["email"] not in existing_emails:
                add_lead(db, {
                    "company":        r["company"],
                    "contact":        r["contact"],
                    "email":          r["email"],
                    "original_email": f"Subject: {r['subject']}\n\n{r['body']}",
                    "sent_date":      sent_date,
                })
                added += 1
        if added:
            save_db(db)
            print(f"  📋 已自动录入跟进数据库：{added} 条 → leads_db.json")
    except ImportError:
        print("  ℹ️  steel_followup.py 不在同目录，跳过跟进同步")

# ══════════════════════════════════════════
# 快捷功能：把 steel_master.py 的 email_results 直接发送
# ══════════════════════════════════════════
def send_from_master_results(email_results: list, dry_run: bool = False) -> dict:
    """
    供 steel_master.py 直接调用：
        from steel_sender import send_from_master_results
        send_from_master_results(email_results)

    dry_run=True：只打印不实际发送（测试用）
    """
    success_count = 0
    fail_count    = 0
    today_str     = str(date.today())

    for item in email_results:
        original = item["original"]
        email_text = item["email"]
        to_address = original.get("email", "")
        contact    = original.get("contact", "")
        company    = original.get("company", "")

        if not to_address or "@" not in to_address:
            print(f"  ⚠️  {company} 无邮箱地址，跳过")
            fail_count += 1
            continue

        subject, body = parse_email_text(email_text)
        if not subject:
            subject = f"Steel Pipe Supply Partnership — {company}"

        if dry_run:
            print(f"  [DRY RUN] {company} → {to_address} | {subject[:40]}")
            success_count += 1
            continue

        print(f"  发送 → {company} ({to_address})...", end=" ")
        result = send_single_email(to_address, subject, body, to_alias=contact)
        if result["success"]:
            print("✅")
            success_count += 1
        else:
            print(f"❌ {result.get('error','')[:50]}")
            fail_count += 1

        append_log({
            "date":    today_str,
            "time":    datetime.now().strftime("%H:%M:%S"),
            "company": company,
            "contact": contact,
            "email":   to_address,
            "subject": subject,
            "status":  "success" if result.get("success") else "failed",
            "error":   result.get("error", ""),
        })

        time.sleep(SEND_INTERVAL)

    return {"success": success_count, "failed": fail_count}

# ══════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Abter Steel 邮件发送模块")
    parser.add_argument(
        "--mode",
        choices=["test", "send", "status"],
        default="test",
        help="test=发测试邮件 | send=批量发送CSV | status=查统计",
    )
    args = parser.parse_args()

    print("=" * 55)
    print("  Abter Steel — 阿里云邮件推送模块")
    print(f"  发信地址：{FROM_ADDRESS}  |  别名：{FROM_ALIAS}")
    print(f"  模式：{args.mode.upper()}")
    print("=" * 55)

    if not ACCESS_KEY_SECRET:
        print("\n  ❌ 请先在 .env 中配置 ALIYUN_ACCESS_KEY_SECRET")
        exit(1)

    {"test": mode_test, "send": mode_send, "status": mode_status}[args.mode]()