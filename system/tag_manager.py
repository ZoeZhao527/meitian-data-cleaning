"""
打标管理系统 - 手动标签管理与自动分析
"""
import json, os, csv, io
from datetime import datetime, timedelta
from pathlib import Path

_BASE = Path(__file__).parent.resolve()
_TAGS_FILE = _BASE / "tags.json"

# ── 标签数据读写 ──

def load_tags():
    if _TAGS_FILE.exists():
        try:
            return json.loads(_TAGS_FILE.read_text(encoding="utf-8"))
        except:
            pass
    return {"care_needs": ["补水保湿","抗衰紧致","美白淡斑","敏感修护",
                           "祛痘控油","眼部护理","身体管理","深层清洁",
                           "舒缓修复","肌底活化"], "customer_tags": {}, "auto_tags": {}}

def save_tags(data):
    _TAGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ── 获取客户列表（从订单数据中提取） ──

def get_all_customers():
    """从订单文件提取所有客户唯一列表"""
    import analyzer
    customers = {}
    acq = analyzer.load_acquisition()
    for row in acq:
        cid = str(row.get("客户账号", ""))
        if cid:
            customers[cid] = {
                "id": cid,
                "name": row.get("客户姓名", ""),
                "phone": row.get("手机号", ""),
                "acq_date": str(row.get("纳新日期", "")),
                "order_date": str(row.get("下单时间", "")),
                "channel": row.get("纳新渠道", ""),
                "source": "纳新"
            }
    orders = analyzer.load_detail_orders()
    for row in orders:
        cid = str(row.get("客户账号", ""))
        if cid and cid not in customers:
            customers[cid] = {
                "id": cid,
                "name": row.get("客户姓名", ""),
                "phone": row.get("手机号", ""),
                "acq_date": "",
                "order_date": str(row.get("下单时间", "")),
                "channel": "",
                "source": "订单"
            }
    return list(customers.values())

# ── 智能分析：添加后未购买人群 ──

def analyze_never_purchased():
    """分析已添加但从未购买/购买次数少的客户，按时间周期分组"""
    import analyzer
    from datetime import datetime
    now = datetime.now()

    # 获取所有客户（从纳新数据）
    acq = analyzer.load_acquisition()
    result = {"1周": [], "半个月": [], "1个月": []}

    for row in acq:
        cid = str(row.get("客户账号", ""))
        order_date_str = str(row.get("下单时间", "") or "")
        acq_date_str = str(row.get("纳新日期", "") or row.get("下单时间", ""))

        # 如果没下单时间 = 从未购买
        has_order = bool(order_date_str.strip())

        # 计算添加时间
        acq_date = None
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"]:
            try:
                acq_date = datetime.strptime(acq_date_str.strip(), fmt)
                break
            except:
                continue
        if not acq_date:
            continue

        days_since_acq = (now - acq_date).days
        customer = {
            "id": cid,
            "name": row.get("客户姓名", ""),
            "phone": row.get("手机号", ""),
            "acq_date": acq_date_str,
            "has_order": has_order,
            "days_since_acq": days_since_acq,
            "channel": row.get("纳新渠道", ""),
        }

        if not has_order:
            if days_since_acq >= 7 and days_since_acq < 15:
                result["1周"].append(customer)
            elif days_since_acq >= 15 and days_since_acq < 30:
                result["半个月"].append(customer)
            elif days_since_acq >= 30:
                result["1个月"].append(customer)

    return result

# ── 打标操作 ──

def get_care_needs():
    return load_tags().get("care_needs", [])

def get_tagged_customers():
    """获取已打标客户列表"""
    data = load_tags()
    return data.get("customer_tags", {})

def get_untagged_customers(search=""):
    """获取未打标客户列表"""
    data = load_tags()
    tagged = set(data.get("customer_tags", {}).keys())
    all_customers = get_all_customers()
    untagged = [c for c in all_customers if c["id"] not in tagged]
    if search:
        search = search.lower()
        untagged = [c for c in untagged if search in c["id"].lower()
                     or search in c["name"].lower()
                     or search in c["phone"]]
    return untagged[:200]

def assign_tag(customer_id, tag):
    """给客户打标"""
    data = load_tags()
    customer_tags = data.setdefault("customer_tags", {})
    if customer_id not in customer_tags:
        # 查找客户姓名
        for c in get_all_customers():
            if c["id"] == customer_id:
                customer_tags[customer_id] = {"name": c["name"], "tags": []}
                break
        else:
            customer_tags[customer_id] = {"name": customer_id, "tags": []}
    if tag not in customer_tags[customer_id]["tags"]:
        customer_tags[customer_id]["tags"].append(tag)
    save_tags(data)
    return True

def remove_tag(customer_id, tag):
    """移除客户的标签"""
    data = load_tags()
    customer_tags = data.setdefault("customer_tags", {})
    if customer_id in customer_tags and tag in customer_tags[customer_id]["tags"]:
        customer_tags[customer_id]["tags"].remove(tag)
        if not customer_tags[customer_id]["tags"]:
            del customer_tags[customer_id]
    save_tags(data)
    return True
