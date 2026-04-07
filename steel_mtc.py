# -*- coding: utf-8 -*-
"""
Abter Steel — MTC材质证明书自动生成  v2.0
═════════════════════════════════════════
读取订单JSON → 修复模板命名空间 → 填写MTC → 导出Excel

使用方式：
  python steel_mtc.py --demo              # 用示例数据生成演示MTC
  python steel_mtc.py --template          # 生成空白订单JSON模板
  python steel_mtc.py --order xxx.json    # 填写指定订单
"""

import os
import json
import argparse
import zipfile
import io
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

# ══════════════════════════════════════════
# 配置
# ══════════════════════════════════════════
TEMPLATE_FILE = "A53_SMLS_MTC.xlsx"

# ══════════════════════════════════════════
# 字段映射（根据MTC截图精确定位）
# ══════════════════════════════════════════
HEADER_MAP = {
    "product":    "C7",    # C7:K8 主格子
    "spec":       "N7",    # N7:R8 主格子（原I7错误）
    "license_no": "V7",    # V7:Z8 主格子（原P7错误）
    "contract_no":"AD7",   # AD7:AI8 主格子（原AH7错误）
    "size":       "D9",    # D9:I10 主格子（原C9错误）
    "grade":      "J9",    # J9:K9 主格子（原I9错误）
    "type_of_end":"L9",    # L9:M10 主格子 ✅
    "processing": "Q9",    # Q9:S10 主格子（原R9错误）
    "delivery":   "T9",    # T9:W9 主格子（原V9错误）
    "auditor":    "A43",   # A43:H43 主格子（原X43错误）
    "certifier":  "AH42",  # AH42:AI43 主格子（原AH43错误）
}

MECH_START_ROW = 16
MECH_COLS = {
    "no": "A", "size": "B", "heat_no": "C", "batch_no": "E",
    "gauge": "H", "ys": "L", "ts": "N", "el": "P", "reduction": "R",
}

CHEM_START_ROW = 29
CHEM_COLS = {
    "no": "A", "C": "C", "Mn": "D", "Si": "E", "S": "F", "P": "G",
    "Cr": "H", "Ni": "I", "Cu": "J", "V": "K",
    "pcs": "AF", "wt_kg": "AH", "wt_t": "AI",
}

TEST_ROW = 38
TEST_COLS = {
    "flatten":   "C",    # C38:E38 主格子（原B38错误）
    "cold_bend": "F",    # F38:G38 主格子 ✅
    "flaring":   "H",    # H38:I38 主格子（原I38错误）
    "drift":     "J",    # J38:L38 主格子（原K38错误）
    "ut":        "M",    # M38:O38 主格子 ✅
    "et":        "P",    # P38:R38 主格子 ✅
    "mpi":       "S",    # S38:U38 主格子 ✅
    "flt":       "V",    # V38:X38 主格子（原U38错误）
    "hp":        "Y",    # Y38:Z38 主格子（原W38错误）
    "surface":   "AA",   # AA38:AC38 主格子（原Y38错误，但与summary重叠，需确认）
}

SUMMARY_CELLS = {
    "total_bundles": "AA38",   # AA38:AC38 主格子 ✅
    "total_pcs":     "AD38",   # AD38:AG38 主格子 ✅
    "total_wt":      "AH38",   # AH38:AI38 主格子 ✅
}

# ══════════════════════════════════════════
# 命名空间修复（兼容WPS/旧版Excel生成的文件）
# ══════════════════════════════════════════
def _fix_namespace(src_path: str, dst_path: str):
    """修复非标准ooxml命名空间，使openpyxl能正确读写"""
    NS_MAP = [
        (b'http://purl.oclc.org/ooxml/spreadsheetml/main',
         b'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),
        (b'http://purl.oclc.org/ooxml/officeDocument/relationships',
         b'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
        (b'http://purl.oclc.org/ooxml/drawingml/main',
         b'http://schemas.openxmlformats.org/drawingml/2006/main'),
        (b'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing',
         b'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'),
    ]
    with open(src_path, 'rb') as f:
        data = f.read()

    zf_in  = zipfile.ZipFile(io.BytesIO(data), 'r')
    buf    = io.BytesIO()
    zf_out = zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED)
    for name in zf_in.namelist():
        content = zf_in.read(name)
        if name.endswith('.xml') or name.endswith('.rels'):
            for old, new in NS_MAP:
                content = content.replace(old, new)
        zf_out.writestr(name, content)
    zf_out.close()

    with open(dst_path, 'wb') as f:
        f.write(buf.getvalue())


# ══════════════════════════════════════════
# 单元格写入工具
# ══════════════════════════════════════════
def _write_cell(ws, coord: str, value, align: str = "center"):
    """写入单元格，保留原有样式"""
    try:
        cell = ws[coord]
        cell.value = value
        if align:
            cell.alignment = Alignment(
                horizontal=align,
                vertical="center",
                wrap_text=True
            )
    except Exception as e:
        print(f"  ⚠️  写入 {coord} 失败：{e}")


# ══════════════════════════════════════════
# 核心填充函数
# ══════════════════════════════════════════
def fill_mtc(order: dict, template_path: str = TEMPLATE_FILE) -> str:
    if not Path(template_path).exists():
        print(f"  ❌ 找不到模板文件：{template_path}")
        return ""

    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    contract_no = order.get("header", {}).get("contract_no", "NEW")
    output_path = f"MTC_{contract_no}_{timestamp}.xlsx"

    # Step1: 修复命名空间
    print("  🔧 修复模板命名空间...")
    _fix_namespace(template_path, output_path)

    # Step2: 加载修复后的文件
    wb = openpyxl.load_workbook(output_path)
    if not wb.worksheets:
        print("  ❌ 模板文件无法解析，请用微软Excel重新另存为xlsx格式")
        return ""

    ws = wb.worksheets[0]
    ws.sheet_state = "visible"

    header  = order.get("header", {})
    items   = order.get("items", [])
    tests   = order.get("tests", {})
    summary = order.get("summary", {})

    # ── 头部信息 ──────────────────────────
    print("  📝 填充头部信息...")
    for field, coord in HEADER_MAP.items():
        value = header.get(field, "")
        if value:
            _write_cell(ws, coord, value)

    # ── 力学性能 ──────────────────────────
    print(f"  📝 填充力学性能（{len(items)} 条）...")
    for i, item in enumerate(items[:7]):
        row = MECH_START_ROW + i
        _write_cell(ws, f"{MECH_COLS['no']}{row}",       i + 1)
        _write_cell(ws, f"{MECH_COLS['size']}{row}",     item.get("size", ""))
        _write_cell(ws, f"{MECH_COLS['heat_no']}{row}",  item.get("heat_no", ""))
        _write_cell(ws, f"{MECH_COLS['batch_no']}{row}", item.get("batch_no", ""))
        _write_cell(ws, f"{MECH_COLS['gauge']}{row}",    item.get("gauge", "---"))
        _write_cell(ws, f"{MECH_COLS['ys']}{row}",       item.get("ys", ""))
        _write_cell(ws, f"{MECH_COLS['ts']}{row}",       item.get("ts", ""))
        _write_cell(ws, f"{MECH_COLS['el']}{row}",       item.get("el", ""))
        _write_cell(ws, f"{MECH_COLS['reduction']}{row}",item.get("reduction", "---"))

    # ── 化学成分 ──────────────────────────
    print(f"  📝 填充化学成分（{len(items)} 条）...")
    total_pcs = 0
    total_wt  = 0.0
    for i, item in enumerate(items[:7]):
        row = CHEM_START_ROW + i
        _write_cell(ws, f"{CHEM_COLS['no']}{row}", i + 1)
        for elem in ["C", "Mn", "Si", "S", "P", "Cr", "Ni", "Cu", "V"]:
            val = item.get(elem, "")
            if val != "":
                _write_cell(ws, f"{CHEM_COLS[elem]}{row}", val)

        pcs   = item.get("pcs", 0)
        wt_t  = item.get("wt_t", 0.0)
        wt_kg = round(wt_t * 1000, 0) if wt_t else item.get("wt_kg", 0)
        _write_cell(ws, f"{CHEM_COLS['pcs']}{row}",   pcs)
        _write_cell(ws, f"{CHEM_COLS['wt_kg']}{row}", wt_kg)
        _write_cell(ws, f"{CHEM_COLS['wt_t']}{row}",  wt_t)
        total_pcs += pcs
        total_wt  += wt_t

    # ── 理化试验 ──────────────────────────
    print("  📝 填充理化试验...")
    for field, col in TEST_COLS.items():
        val = tests.get(field, "")
        if val:
            _write_cell(ws, f"{col}{TEST_ROW}", val)

    # ── 汇总 ──────────────────────────────
    print("  📝 填充汇总数据...")
    _write_cell(ws, SUMMARY_CELLS["total_bundles"], summary.get("total_bundles", ""))
    _write_cell(ws, SUMMARY_CELLS["total_pcs"],     summary.get("total_pcs", total_pcs))
    _write_cell(ws, SUMMARY_CELLS["total_wt"],      round(summary.get("total_wt", total_wt), 3))

    wb.save(output_path)
    print(f"\n  ✅ MTC已生成 → {output_path}")
    return output_path


# ══════════════════════════════════════════
# 示例订单（与截图完全一致）
# ══════════════════════════════════════════
DEMO_ORDER = {
    "header": {
        "product":    "SMLS STEEL PIPE",
        "spec":       "ASTM A53",
        "license_no": "MTC109508TGS",
        "contract_no":"AB109508TGS",
        "size":       "AS BELOW",
        "grade":      "GRB",
        "type_of_end":"PE/BE",
        "processing": "Hot rolled",
        "delivery":   "---",
        "auditor":    "董东",
        "certifier":  "易几",
    },
    "items": [
        {"size":"33.4*3.38*6000","heat_no":"19051733","batch_no":"19067",
         "ys":278,"ts":467,"el":33.5,
         "C":0.22,"Mn":0.55,"Si":0.26,"S":0.005,"P":0.012,"Cr":0.12,"Ni":0.04,"Cu":0.01,"V":0.01,
         "pcs":11,"wt_t":20.260},
        {"size":"42.20*3.56*6000","heat_no":"19052265","batch_no":"19068",
         "ys":286,"ts":466,"el":31.0,
         "C":0.23,"Mn":0.56,"Si":0.27,"S":0.006,"P":0.017,"Cr":0.08,"Ni":0.05,"Cu":0.02,"V":0.01,
         "pcs":6,"wt_t":10.290},
        {"size":"60.3*3.91*6000","heat_no":"19067751","batch_no":"19069",
         "ys":285,"ts":475,"el":32.5,
         "C":0.22,"Mn":0.58,"Si":0.27,"S":0.005,"P":0.017,"Cr":0.08,"Ni":0.05,"Cu":0.02,"V":0.01,
         "pcs":3,"wt_t":5.324},
        {"size":"73.0*5.16*6000","heat_no":"19051169","batch_no":"19070",
         "ys":296,"ts":462,"el":34.0,
         "C":0.21,"Mn":0.60,"Si":0.25,"S":0.008,"P":0.015,"Cr":0.06,"Ni":0.05,"Cu":0.01,"V":0.01,
         "pcs":2,"wt_t":3.650},
        {"size":"60.3*3*6000","heat_no":"19051162","batch_no":"19071",
         "ys":290,"ts":473,"el":31.5,
         "C":0.25,"Mn":0.55,"Si":0.27,"S":0.006,"P":0.016,"Cr":0.05,"Ni":0.04,"Cu":0.02,"V":0.01,
         "pcs":8,"wt_t":12.307},
        {"size":"88.9*4*6000","heat_no":"19064566","batch_no":"19072",
         "ys":302,"ts":475,"el":33.0,
         "C":0.24,"Mn":0.49,"Si":0.26,"S":0.009,"P":0.013,"Cr":0.06,"Ni":0.03,"Cu":0.01,"V":0.01,
         "pcs":17,"wt_t":19.635},
        {"size":"88.9*5.49*6000","heat_no":"19068857","batch_no":"19073",
         "ys":295,"ts":469,"el":33.5,
         "C":0.20,"Mn":0.52,"Si":0.27,"S":0.007,"P":0.012,"Cr":0.04,"Ni":0.05,"Cu":0.02,"V":0.01,
         "pcs":17,"wt_t":26.590},
    ],
    "tests": {
        "flatten": "Accepted", "cold_bend": "---",  "flaring": "Accepted",
        "drift":   "---",      "ut": "Accepted",    "et": "Accepted",
        "mpi":     "---",      "flt": "---",        "hp": "Accepted",
        "surface": "Accepted",
    },
    "summary": {
        "total_bundles": 64,
        "total_pcs":     3461,
        "total_wt":      98.056,
    }
}


# ══════════════════════════════════════════
# 空白订单模板
# ══════════════════════════════════════════
def save_order_template():
    template = {
        "header": {
            "product":    "SMLS STEEL PIPE",
            "spec":       "ASTM A53",
            "license_no": "填写证书编号",
            "contract_no":"填写合同编号",
            "size":       "AS BELOW",
            "grade":      "GRB",
            "type_of_end":"PE/BE",
            "processing": "Hot rolled",
            "delivery":   "---",
            "auditor":    "填写审核人姓名",
            "certifier":  "填写制证人姓名",
        },
        "items": [
            {
                "size": "外径*壁厚*长度 例:33.4*3.38*6000",
                "heat_no": "炉号", "batch_no": "批号",
                "ys": 0, "ts": 0, "el": 0.0, "reduction": "---",
                "C":0.0,"Mn":0.0,"Si":0.0,"S":0.0,"P":0.0,
                "Cr":0.0,"Ni":0.0,"Cu":0.0,"V":0.0,
                "pcs": 0, "wt_t": 0.0,
            }
        ],
        "tests": {
            "flatten": "Accepted", "cold_bend": "---",
            "flaring": "Accepted", "drift": "---",
            "ut": "Accepted",      "et": "Accepted",
            "mpi": "---",          "flt": "---",
            "hp": "Accepted",      "surface": "Accepted",
        },
        "summary": {"total_bundles": 0, "total_pcs": 0, "total_wt": 0.0}
    }
    with open("order_template.json", "w", encoding="utf-8") as f:
        json.dump(template, f, ensure_ascii=False, indent=2)
    print("  ✅ 空白模板 → order_template.json")
    print("  📝 填写后运行：python steel_mtc.py --order order_template.json")


# ══════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Abter Steel MTC生成系统 v2.0")
    parser.add_argument("--demo",     action="store_true", help="用示例数据生成演示MTC")
    parser.add_argument("--order",    type=str,            help="订单JSON文件路径")
    parser.add_argument("--template", action="store_true", help="生成空白订单模板JSON")
    args = parser.parse_args()

    print("=" * 55)
    print("  Abter Steel — MTC材质证明书自动生成  v2.0")
    print("=" * 55)

    if args.template:
        save_order_template()
    elif args.demo:
        print("\n  📋 演示模式\n")
        output = fill_mtc(DEMO_ORDER)
        if output:
            print(f"\n  📌 用Excel打开 {output} 查看结果")
            print(f"  ⚠️  如有字段位置偏差，告诉我，我来修正")
    elif args.order:
        if not Path(args.order).exists():
            print(f"  ❌ 找不到：{args.order}")
        else:
            with open(args.order, "r", encoding="utf-8") as f:
                order = json.load(f)
            fill_mtc(order)
    else:
        print("\n  用法：")
        print("  python steel_mtc.py --demo")
        print("  python steel_mtc.py --template")
        print("  python steel_mtc.py --order xxx.json")