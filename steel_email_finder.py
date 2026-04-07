# -*- coding: utf-8 -*-
"""
Abter Steel — 邮箱挖掘模块  v1.0
══════════════════════════════════
三层挖掘策略（全部基于 Serper，零额外成本）：
  第1层：直接搜索公司域名邮箱
  第2层：搜索公司官网 → 提取 Contact/About 页面邮箱
  第3层：搜索 LinkedIn/公司新闻里的联系人邮箱

后期升级：
  把 _apollo_find() 解注释，填入 Apollo Key，准确率从 35% → 90%

运行方式：
  python steel_email_finder.py --mode find    # 从 leads_db.json 挖邮箱
  python steel_email_finder.py --mode test    # 测试单个公司
  python steel_email_finder.py --mode enrich  # 批量补全已有CSV的邮箱列

对接方式（在 steel_master.py 末尾加）：
  from steel_email_finder import enrich_leads
  email_results = enrich_leads(email_results)
"""

import os
import re
import time
import json
import csv
import argparse
from pathlib import Path
from datetime import datetime

import requests
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════
# 配置
# ══════════════════════════════════════════
SERPER_KEY     = os.environ.get("SERPER_API_KEY", "")
SEARCH_DELAY   = 0.8   # 每次搜索间隔（秒），避免限流
MAX_RESULTS    = 10    # 每次搜索返回条数

# 邮箱正则（匹配 xxx@xxx.xxx 格式）
EMAIL_PATTERN  = re.compile(
    r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
    re.IGNORECASE
)

# 过滤掉明显无效的邮箱（图片/样式/示例）
SKIP_DOMAINS = {
    "example.com", "test.com", "email.com", "domain.com",
    "yourdomain.com", "company.com", "sentry.io", "github.com",
    "google.com", "microsoft.com", "adobe.com", "w3.org",
    "schema.org", "placeholder.com",
}

SKIP_PREFIXES = {
    "noreply", "no-reply", "donotreply", "mailer", "daemon",
    "postmaster", "webmaster", "info@info", "admin@admin",
}

# 采购相关关键词（优先级排序）
PROCUREMENT_KEYWORDS = [
    "procurement", "purchasing", "supply", "sourcing",
    "buyer", "purchase", "material", "vendor",
]

# ══════════════════════════════════════════
# Serper 搜索工具
# ══════════════════════════════════════════
def _serper_search(query: str, num: int = 10) -> list:
    if not SERPER_KEY:
        print("  ⚠️  SERPER_API_KEY 未配置")
        return []
    try:
        resp = requests.post(
            "https://google.serper.dev/search",
            headers={"X-API-KEY": SERPER_KEY, "Content-Type": "application/json"},
            json={"q": query, "num": num},
            timeout=15,
        )
        return resp.json().get("organic", [])
    except Exception as e:
        print(f"  ⚠️  搜索失败：{e}")
        return []

def _extract_emails_from_text(text: str) -> list:
    """从文本中提取所有有效邮箱"""
    found = EMAIL_PATTERN.findall(text)
    valid = []
    for email in found:
        email = email.lower().strip(".,;:")
        domain = email.split("@")[-1]
        prefix = email.split("@")[0]
        if domain in SKIP_DOMAINS:
            continue
        if any(prefix.startswith(p) for p in SKIP_PREFIXES):
            continue
        if len(email) > 100:
            continue
        if email not in valid:
            valid.append(email)
    return valid

def _score_email(email: str, company: str = "", title: str = "") -> int:
    """
    给邮箱打分，分越高越可能是采购联系人
    最高 100 分
    """
    score = 50  # 基础分
    prefix = email.split("@")[0].lower()
    domain  = email.split("@")[-1].lower()

    # 采购关键词加分
    for kw in PROCUREMENT_KEYWORDS:
        if kw in prefix:
            score += 25
            break

    # 公司名匹配加分
    if company:
        company_clean = re.sub(r'[^a-z0-9]', '', company.lower())
        domain_clean  = re.sub(r'[^a-z0-9]', '', domain)
        if company_clean[:6] in domain_clean or domain_clean[:6] in company_clean:
            score += 15

    # 职位关键词匹配加分
    if title:
        for kw in PROCUREMENT_KEYWORDS:
            if kw in title.lower() and kw in prefix:
                score += 10
                break

    # 通用邮箱减分
    generic_prefixes = ["info", "contact", "sales", "admin", "support", "hello"]
    if prefix in generic_prefixes:
        score -= 10

    return min(score, 100)

# ══════════════════════════════════════════
# 三层挖掘策略
# ══════════════════════════════════════════
def _layer1_direct_search(company: str, country: str) -> list:
    """
    第1层：直接搜索公司采购邮箱
    最准确，命中率约 20-30%
    """
    emails = []
    queries = [
        f'"{company}" procurement email contact',
        f'"{company}" "procurement@" OR "purchasing@" OR "supply@"',
        f'"{company}" {country} purchasing manager email',
    ]
    for q in queries:
        results = _serper_search(q, num=8)
        for r in results:
            text = r.get("title", "") + " " + r.get("snippet", "")
            emails.extend(_extract_emails_from_text(text))
        time.sleep(SEARCH_DELAY)
    return list(set(emails))


def _layer2_website_search(company: str, website: str = "") -> list:
    """
    第2层：搜索公司官网的 Contact/About 页面
    命中率约 30-40%（通用邮箱为主）
    """
    emails = []

    # 如果有官网，直接搜官网联系页
    if website:
        domain = website.replace("https://", "").replace("http://", "").split("/")[0]
        queries = [
            f'site:{domain} contact email procurement',
            f'site:{domain} "email" OR "@{domain}"',
        ]
    else:
        queries = [
            f'"{company}" official website contact email',
            f'"{company}" contact us email address',
        ]

    for q in queries:
        results = _serper_search(q, num=8)
        for r in results:
            text = r.get("title", "") + " " + r.get("snippet", "") + " " + r.get("link", "")
            emails.extend(_extract_emails_from_text(text))
        time.sleep(SEARCH_DELAY)
    return list(set(emails))


def _layer3_linkedin_news_search(company: str, title: str = "") -> list:
    """
    第3层：从 LinkedIn 和新闻稿里找邮箱
    命中率约 10-15%，但质量高（真实个人邮箱）
    """
    emails = []
    role = title or "procurement manager"
    queries = [
        f'site:linkedin.com "{company}" "{role}" email',
        f'"{company}" "{role}" email contact 2024 OR 2025',
        f'"{company}" press release contact email',
    ]
    for q in queries:
        results = _serper_search(q, num=5)
        for r in results:
            text = r.get("title", "") + " " + r.get("snippet", "")
            emails.extend(_extract_emails_from_text(text))
        time.sleep(SEARCH_DELAY)
    return list(set(emails))


# ══════════════════════════════════════════
# Apollo API（后期升级用，现在注释掉）
# ══════════════════════════════════════════
# def _apollo_find(company: str, title: str = "") -> list:
#     """付费升级：Apollo.io API，准确率 90%+"""
#     APOLLO_KEY = os.environ.get("APOLLO_API_KEY", "")
#     resp = requests.post(
#         "https://api.apollo.io/v1/people/search",
#         headers={"Content-Type": "application/json", "Cache-Control": "no-cache"},
#         json={
#             "api_key": APOLLO_KEY,
#             "q_organization_name": company,
#             "person_titles": [title or "procurement manager", "purchasing manager", "supply chain manager"],
#             "per_page": 5,
#         }
#     )
#     people = resp.json().get("people", [])
#     emails = [p.get("email") for p in people if p.get("email")]
#     return emails


# ══════════════════════════════════════════
# 主查找函数
# ══════════════════════════════════════════
def find_email(company: str, country: str = "", title: str = "",
               website: str = "", verbose: bool = True) -> dict:
    """
    对一家公司执行三层挖掘，返回最优邮箱
    返回：{
        "email": "best@company.com",
        "all_emails": ["a@b.com", ...],
        "confidence": 85,
        "source": "layer1",
        "company": "...",
    }
    """
    if verbose:
        print(f"  🔍 挖掘：{company}...")

    all_emails = []
    source     = "none"

    # 第1层
    layer1 = _layer1_direct_search(company, country)
    if layer1:
        all_emails.extend(layer1)
        source = "layer1_direct"
        if verbose:
            print(f"     第1层找到：{layer1}")

    # 第2层（无论第1层有没有找到，都补充）
    layer2 = _layer2_website_search(company, website)
    new2   = [e for e in layer2 if e not in all_emails]
    if new2:
        all_emails.extend(new2)
        if not layer1:
            source = "layer2_website"
        if verbose:
            print(f"     第2层找到：{new2}")

    # 第3层（只在前两层没找到时才用，节省搜索次数）
    if not all_emails:
        layer3 = _layer3_linkedin_news_search(company, title)
        new3   = [e for e in layer3 if e not in all_emails]
        if new3:
            all_emails.extend(new3)
            source = "layer3_linkedin"
            if verbose:
                print(f"     第3层找到：{new3}")

    if not all_emails:
        if verbose:
            print(f"     ❌ 未找到邮箱")
        return {
            "email":      "",
            "all_emails": [],
            "confidence": 0,
            "source":     "not_found",
            "company":    company,
        }

    # 打分排序，选最优邮箱
    scored = [(e, _score_email(e, company, title)) for e in all_emails]
    scored.sort(key=lambda x: x[1], reverse=True)
    best_email, best_score = scored[0]

    if verbose:
        print(f"     ✅ 最优邮箱：{best_email}（置信度 {best_score}%）")

    return {
        "email":      best_email,
        "all_emails": [e for e, _ in scored],
        "confidence": best_score,
        "source":     source,
        "company":    company,
    }

# ══════════════════════════════════════════
# 批量对接：给 steel_master.py 的 email_results 补充邮箱
# ══════════════════════════════════════════
def enrich_leads(email_results: list) -> list:
    """
    在 steel_master.py 末尾调用：
        from steel_email_finder import enrich_leads
        email_results = enrich_leads(email_results)

    自动给每个 item["original"] 补充 email 字段
    """
    print(f"\n📧 邮箱挖掘开始（{len(email_results)} 家客户）...\n")
    found = 0

    for item in email_results:
        original = item["original"]
        company  = original.get("company", "")
        country  = original.get("country", "")
        title    = original.get("title", "")

        # 已有邮箱则跳过
        if original.get("email") and "@" in original.get("email", ""):
            print(f"  ✓ {company} 已有邮箱：{original['email']}")
            continue

        result = find_email(company, country, title)
        if result["email"]:
            original["email"]            = result["email"]
            original["email_confidence"] = result["confidence"]
            original["email_source"]     = result["source"]
            original["email_all"]        = result["all_emails"]
            found += 1
        else:
            original["email"]            = ""
            original["email_confidence"] = 0
            original["email_source"]     = "not_found"

        time.sleep(0.5)

    print(f"\n  📧 邮箱挖掘完成：{found}/{len(email_results)} 家找到邮箱")
    return email_results

# ══════════════════════════════════════════
# 补全已有 CSV 文件的邮箱列
# ══════════════════════════════════════════
def enrich_csv(input_csv: str, output_csv: str = ""):
    """
    读取已有 CSV（含公司名），补充邮箱列，输出新 CSV
    """
    if not output_csv:
        stem       = Path(input_csv).stem
        output_csv = f"{stem}_with_emails.csv"

    rows = []
    with open(input_csv, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        rows = list(reader)

    print(f"\n📧 批量补全邮箱：{len(rows)} 条记录\n")

    # 添加邮箱相关列
    new_fields = ["挖掘邮箱", "邮箱置信度", "邮箱来源"]
    all_fields = list(fieldnames) + [f for f in new_fields if f not in fieldnames]

    found = 0
    for i, row in enumerate(rows, 1):
        company = row.get("公司") or row.get("company") or row.get("Company", "")
        country = row.get("国家") or row.get("country", "")
        title   = row.get("职位") or row.get("title", "")
        existing_email = row.get("挖掘邮箱") or row.get("邮箱", "")

        if existing_email and "@" in existing_email:
            print(f"  [{i:03d}] {company} 已有邮箱，跳过")
            continue

        print(f"  [{i:03d}/{len(rows)}] {company}...")
        result = find_email(company, country, title, verbose=False)

        row["挖掘邮箱"]   = result["email"]
        row["邮箱置信度"] = f"{result['confidence']}%"
        row["邮箱来源"]   = result["source"]

        if result["email"]:
            found += 1
            print(f"  [{i:03d}] ✅ {company} → {result['email']} ({result['confidence']}%)")
        else:
            print(f"  [{i:03d}] ❌ {company} → 未找到")

        time.sleep(0.5)

    # 写出新 CSV
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=all_fields)
        writer.writeheader()
        writer.writerows(rows)

    print(f"\n  ✅ 完成：{found}/{len(rows)} 家找到邮箱")
    print(f"  📄 输出文件：{output_csv}")
    return output_csv

# ══════════════════════════════════════════
# 模式：test — 测试单个公司
# ══════════════════════════════════════════
def mode_test():
    print("\n🧪 单公司邮箱测试\n")
    company = input("  公司名：").strip()
    country = input("  国家（可选）：").strip()
    title   = input("  职位关键词（可选，如 procurement manager）：").strip()
    website = input("  官网（可选）：").strip()

    print()
    result = find_email(company, country, title, website, verbose=True)

    print(f"\n  {'═'*50}")
    print(f"  公司：{company}")
    print(f"  最优邮箱：{result['email'] or '未找到'}")
    print(f"  置信度：{result['confidence']}%")
    print(f"  来源：{result['source']}")
    if len(result["all_emails"]) > 1:
        print(f"  其他候选：{', '.join(result['all_emails'][1:4])}")
    print(f"  {'═'*50}\n")

# ══════════════════════════════════════════
# 模式：find — 从 leads_db.json 批量挖掘
# ══════════════════════════════════════════
def mode_find():
    print("\n📧 从 leads_db.json 批量挖掘邮箱\n")

    db_file = Path("leads_db.json")
    if not db_file.exists():
        print("  ⚠️  leads_db.json 不存在，请先运行 steel_master.py 生成线索")
        return

    with open(db_file, "r", encoding="utf-8") as f:
        db = json.load(f)

    # 找没有邮箱的线索
    no_email = [l for l in db["leads"] if not l.get("email") or "@" not in l.get("email", "")]
    print(f"  共 {len(db['leads'])} 条线索，其中 {len(no_email)} 条缺少邮箱")

    if not no_email:
        print("  ✅ 所有线索都有邮箱了！")
        return

    confirm = input(f"\n  开始挖掘 {len(no_email)} 条？(y/n) ").strip().lower()
    if confirm != "y":
        return

    found = 0
    for lead in no_email:
        result = find_email(
            company = lead.get("company", ""),
            country = lead.get("country", ""),
            title   = lead.get("title", ""),
        )
        if result["email"]:
            lead["email"]            = result["email"]
            lead["email_confidence"] = result["confidence"]
            lead["email_source"]     = result["source"]
            found += 1
        time.sleep(0.3)

    # 保存回 DB
    with open(db_file, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

    print(f"\n  ✅ 完成：{found}/{len(no_email)} 家找到邮箱")
    print(f"  💾 已更新 leads_db.json")

    # 同时导出 CSV 方便查看
    output_csv = f"leads_with_emails_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["公司", "国家", "联系人", "职位", "邮箱", "置信度", "状态"])
        for lead in db["leads"]:
            writer.writerow([
                lead.get("company", ""),
                lead.get("country", ""),
                lead.get("contact", ""),
                lead.get("title", ""),
                lead.get("email", ""),
                f"{lead.get('email_confidence', 0)}%",
                lead.get("status", ""),
            ])
    print(f"  📄 同步导出：{output_csv}")

# ══════════════════════════════════════════
# 模式：enrich — 补全已有 CSV
# ══════════════════════════════════════════
def mode_enrich():
    print("\n📧 补全 CSV 文件邮箱\n")

    csv_files = sorted(Path(".").glob("*.csv"))
    if not csv_files:
        csv_path = input("  输入CSV文件路径：").strip()
    else:
        print("  找到以下CSV文件：")
        for i, f in enumerate(csv_files, 1):
            print(f"  {i}. {f.name}")
        idx      = int(input("\n  选择哪个（序号）：").strip()) - 1
        csv_path = str(csv_files[idx])

    enrich_csv(csv_path)

# ══════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Abter Steel 邮箱挖掘模块")
    parser.add_argument(
        "--mode",
        choices=["test", "find", "enrich"],
        default="test",
        help="test=测试单个 | find=从DB批量挖 | enrich=补全CSV",
    )
    args = parser.parse_args()

    print("=" * 55)
    print("  Abter Steel — 邮箱挖掘模块  v1.0")
    print(f"  模式：{args.mode.upper()}")
    print(f"  引擎：Serper Search（免费）")
    print("=" * 55)

    if not SERPER_KEY:
        print("\n  ❌ 请先在 .env 中配置 SERPER_API_KEY")
        exit(1)

    {"test": mode_test, "find": mode_find, "enrich": mode_enrich}[args.mode]()