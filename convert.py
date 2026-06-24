"""
EDC Parameter Config — Excel to JSON Converter
วิธีใช้:
  1. วางไฟล์ Excel ทุกแบงค์ไว้ในโฟลเดอร์เดียวกับ convert.py
  2. รัน: python convert.py
  3. จะได้ data.json สำหรับ upload ขึ้น GitHub
"""

import json, os, glob
import pandas as pd
from datetime import datetime

OUTPUT = "data.json"

def safe(v):
    if v is None: return ""
    s = str(v).strip()
    if s in ("None","nan","NaT","NaN",""): return ""
    if s.endswith(".0"):
        try: return str(int(float(s)))
        except: pass
    return s.lstrip("'")

def yn(v):
    if v is None: return False
    return str(v).strip().upper() in ("Y","YES","1","TRUE","X","✓")

records = []

# ════════════════════════════════════════
# KBANK
# ════════════════════════════════════════
files = glob.glob("*KBANK*.xlsx")
if files:
    print(f"[KBANK] อ่านไฟล์: {files[0]}")
    df = pd.read_excel(files[0], dtype=str, sheet_name="ParameterReport")
    for _, r in df.iterrows():
        tid = safe(r.get("TID",""))
        if not tid or tid == "0": continue
        hosts = {}
        for col in df.columns:
            if col.startswith("TID_") or col.startswith("MID_"):
                v = safe(r.get(col,""))
                if v and v not in ("0","-"): hosts[col] = v
        records.append({
            "bank":"KBANK", "tid":tid,
            "mid":safe(r.get("MID","")),
            "sn":safe(r.get("S/N EDC","")),
            "slip1":safe(r.get("Slip Line_1","")),
            "slip2":safe(r.get("Slip Line_2","")),
            "slip3":safe(r.get("Slip Line_3","")),
            "model":safe(r.get("Model","")),
            "sw":safe(r.get("SW Version","")),
            "ip":safe(r.get("IP","")),
            "tel_system":safe(r.get("Tel_system","")),
            "apn":safe(r.get("APN SIM","")),
            "flow_wallet":safe(r.get("Flow_Type_Wallet","")),
            "settle_mode":safe(r.get("Settle Mode","")),
            "time_settle":safe(r.get("Time settle","")),
            "control_limit":safe(r.get("จำกัดยอดรูด-swipe limit","")),
            "linkpos":yn(r.get("ERM_Enabled","")),
            "keyin":any(yn(r.get(c,"")) for c in ["Function_KBANK_Key-in.","Function_TPN_Key-in.","Function_AMEX_Key-in.","Function_DCC_Key-in.","Function_Redeem_Key-in."]),
            "tip":any(yn(r.get(c,"")) for c in ["Function_KBANK_Tip.","Function_TPN_Tips.","Function_AMEX_Tip.","Function_DCC_Tip."]),
            "offline":any(yn(r.get(c,"")) for c in ["Function_KBANK_OFFline.","Function_AMEX_OFFline.","Function_DCC_OFFline."]),
            "refund":any(yn(r.get(c,"")) for c in ["Function_KBANK_ReFund.","Function_TPN_ReFund.","Function_AMEX_ReFund."]),
            "cardver":yn(r.get("Exp. Date","")),
            "dynamic_offline":yn(r.get("Dynamic_Offline","")),
            "multi":yn(r.get("Multi_mechant","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len([x for x in records if x['bank']=='KBANK']):,} records")
else:
    print("[KBANK] ไม่พบไฟล์ *KBANK*.xlsx")

# ════════════════════════════════════════
# WOM
# ════════════════════════════════════════
files = glob.glob("*WOM*.xlsx") or glob.glob("*Common_Template*.xlsx")
if files:
    print(f"[WOM] อ่านไฟล์: {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    for _, r in df.iterrows():
        tid = safe(r.get("MAIN_TID",""))
        if not tid or tid == "0": continue
        hosts = {}
        for col in df.columns:
            if col.startswith("TID_") or col.startswith("MID_"):
                v = safe(r.get(col,""))
                if v and v not in ("0","-"): hosts[col] = v
        records.append({
            "bank":"WOM", "tid":tid,
            "mid":safe(r.get("MAIN_MID","")),
            "sn":safe(r.get("SERIAL_NO","")),
            "slip1":safe(r.get("SLIP_LINE_1","")),
            "slip2":safe(r.get("SLIP_LINE_2","")),
            "slip3":safe(r.get("SLIP_LINE_3","")),
            "model":safe(r.get("MODEL","")),
            "sw":safe(r.get("SW_VERSION","")),
            "ip":safe(r.get("IP","")),
            "tel_system":safe(r.get("TEL_SYSTEM","")),
            "apn":safe(r.get("APN_SIM","")),
            "flow_wallet":safe(r.get("FLOW_TYPE_WALLET","")),
            "settle_mode":safe(r.get("SETTLE_MODE","")),
            "time_settle":safe(r.get("TIME_SET","")),
            "control_limit":safe(r.get("LIMIT_AMOUNT","")),
            "linkpos":safe(r.get("COMM_TYPE","")) not in ("","0"),
            "keyin":any(yn(r.get(c,"")) for c in ["FUNCTION_KBANK_KEYIN","FUNCTION_TPN_KEYIN","FUNCTION_AMEX_KEYIN","FUNCTION_DCC_KEYIN","FUNCTION_REDEEM_KEYIN"]),
            "tip":any(yn(r.get(c,"")) for c in ["FUNCTION_KBANK_TIP","FUNCTION_AMEX_TIP","FUNCTION_DCC_TIP","FUNCTION_TPN_TIPS"]),
            "offline":any(yn(r.get(c,"")) for c in ["FUNCTION_KBANK_OFFLINE","FUNCTION_AMEX_OFFLINE","FUNCTION_DCC_OFFLINE"]),
            "refund":any(yn(r.get(c,"")) for c in ["FUNCTION_KBANK_REFUND","FUNCTION_AMEX_REFUND","FUNCTION_TPN_REFUND"]),
            "cardver":yn(r.get("EXP_DATE","")),
            "dynamic_offline":yn(r.get("DYNAMIC_OFFLINE","")),
            "multi":yn(r.get("MULTI_MERCHANT","")),
            "lounge":yn(r.get("LOUNGE_VISIT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len([x for x in records if x['bank']=='WOM']):,} records")
else:
    print("[WOM] ไม่พบไฟล์ *WOM*.xlsx")

# ════════════════════════════════════════
# BAY
# ════════════════════════════════════════
files = glob.glob("*BAY*.xlsx")
if files:
    print(f"[BAY] อ่านไฟล์: {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    for _, r in df.iterrows():
        tid = safe(r.get("TIDหลัก",""))
        if not tid or tid == "0": continue
        hosts = {}
        for col in df.columns:
            if any(x in col for x in ["TID_ACQUIRER","MID_ACQUIRER","TID_ARKE","MID_ARKE","TID_KBANK","MID_KBANK","TID_KTC","MID_KTC"]):
                v = safe(r.get(col,""))
                if v and v not in ("0","-"):
                    acol = col.replace("TID_","ACQUIRER").replace("MID_","ACQUIRER")
                    acq = safe(r.get(acol,""))
                    key = f"{acq}/{col}" if acq else col
                    hosts[key] = v
        records.append({
            "bank":"BAY", "tid":tid,
            "mid":safe(r.get("MID_ACQUIRER1","")),
            "sn":safe(r.get("serialNumber","")),
            "slip1":safe(r.get("SLIP1","")),
            "slip2":safe(r.get("SLIP2","")),
            "slip3":safe(r.get("SLIP3","")),
            "model":safe(r.get("PROFILE","")),
            "sw":"",
            "ip":safe(r.get("IP ทำรายการ Primary","")),
            "tel_system":"",
            "apn":"", "flow_wallet":"",
            "settle_mode":"",
            "time_settle":safe(r.get("AUTO SETTLEMENT TIME","")),
            "control_limit":"",
            "linkpos":yn(r.get("LINK POS","")),
            "keyin":yn(r.get("KEYIN","")),
            "tip":yn(r.get("ADJUST","")),
            "offline":yn(r.get("OFFLINE","")),
            "refund":yn(r.get("REFUND","")),
            "cardver":False, "dynamic_offline":False,
            "multi":yn(r.get("MULTI-MERCHANT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len([x for x in records if x['bank']=='BAY']):,} records")
else:
    print("[BAY] ไม่พบไฟล์ *BAY*.xlsx")

# ════════════════════════════════════════
# BBL
# ════════════════════════════════════════
files = glob.glob("*BBL*.xlsx")
if files:
    print(f"[BBL] อ่านไฟล์: {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    for _, r in df.iterrows():
        tid = safe(r.get("Terminal ID",""))
        if not tid or tid == "0": continue
        hosts = {}
        for col in ["TID UPI","BBL TID","BBL MID","TID-DCC","TID-UPI-DCI","TID-BE-SMART","MID-BE-SMART",
                    "TID-REDEMPTION","MID-REDEMPTION","TID.QR REF1 (TID)","TID_ALIPAY + WECHAT","TID BSS","MID BSS"]:
            v = safe(r.get(col,""))
            if v and v not in ("0","-"): hosts[col] = v
        records.append({
            "bank":"BBL", "tid":tid,
            "mid":safe(r.get("BBL MID","")),
            "sn":safe(r.get("Serial No.","")),
            "slip1":safe(r.get("LINE 1","")),
            "slip2":safe(r.get("LINE 2","")),
            "slip3":safe(r.get("LINE 3","")),
            "model":safe(r.get("Model","")),
            "sw":"",
            "ip":"", "tel_system":"", "apn":"", "flow_wallet":"",
            "settle_mode":safe(r.get("Force Settlement","")),
            "time_settle":safe(r.get("Auto Settlement","")),
            "control_limit":safe(r.get("Transaction Limit","")),
            "linkpos":yn(r.get("LINKPOS","")),
            "keyin":yn(r.get("Manual Key-in","")),
            "tip":yn(r.get("Adj. Tip","")),
            "offline":yn(r.get("Offline Sale","")),
            "refund":yn(r.get("Refund","")),
            "cardver":False,
            "dynamic_offline":yn(r.get("DYNAMIC_OFFLINE","")),
            "multi":yn(r.get("MULTI","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len([x for x in records if x['bank']=='BBL']):,} records")
else:
    print("[BBL] ไม่พบไฟล์ *BBL*.xlsx")

# ════════════════════════════════════════
# SCB
# ════════════════════════════════════════
files = glob.glob("*SCB*.xlsx")
if files:
    print(f"[SCB] อ่านไฟล์: {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    for _, r in df.iterrows():
        tid = safe(r.get("HOST.1/TERMINAL_ID",""))
        if not tid or tid == "0": continue
        hosts = {}
        for i in range(1,13):
            t = safe(r.get(f"HOST.{i}/TERMINAL_ID",""))
            m = safe(r.get(f"HOST.{i}/MERCHANT_ID",""))
            qn = safe(r.get(f"HOST.{i}/QR_MERCHANT_NAME",""))
            if t and t not in ("0","-"): hosts[f"HOST.{i} TID"] = t
            if m and m not in ("0","-"): hosts[f"HOST.{i} MID"] = m
            if qn: hosts[f"HOST.{i} QR_NAME"] = qn
        records.append({
            "bank":"SCB", "tid":tid,
            "mid":safe(r.get("HOST.1/MERCHANT_ID","")),
            "sn":safe(r.get("serialNumber","")),
            "slip1":safe(r.get("PRINT_CONFIG.1/HEADER1","")),
            "slip2":safe(r.get("PRINT_CONFIG.1/HEADER2","")),
            "slip3":safe(r.get("PRINT_CONFIG.1/HEADER3","")),
            "model":safe(r.get("modelName","")),
            "sw":"",
            "ip":safe(r.get("CONNECTION.1/PRIMARY_IP","")),
            "tel_system":safe(r.get("HOST.1/ACTIVE_MEDIA","")),
            "apn":"", "flow_wallet":"",
            "settle_mode":safe(r.get("HOST.1/ENABLE_AUTO_SETTLEMENT","")),
            "time_settle":safe(r.get("HOST.1/AUTO_SETTLE_TIME","")),
            "control_limit":safe(r.get("TERMINAL/MAX_TRANS_AMOUNT","")),
            "linkpos":yn(r.get("TERMINAL/ENABLE_ECR","")),
            "keyin":yn(r.get("TRANS_SWITCH.1/ENABLE_MANUAL_KEY_IN","")),
            "tip":yn(r.get("TERMINAL/ALLOW_TIP","")),
            "offline":False,
            "refund":yn(r.get("MERCHANT.1/ALLOW_CARD_REFUND","")),
            "cardver":False,
            "dynamic_offline":False,
            "multi":yn(r.get("TERMINAL/ENABLE_MULTI_MERCHANT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len([x for x in records if x['bank']=='SCB']):,} records")
else:
    print("[SCB] ไม่พบไฟล์ *SCB*.xlsx")

# ════════════════════════════════════════
# บันทึก data.json
# ════════════════════════════════════════
out = {
    "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "total": len(records),
    "records": records
}
with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(",",":"))

size = os.path.getsize(OUTPUT)/1024/1024
print(f"\n{'='*40}")
print(f"✅ เสร็จแล้ว! {len(records):,} records → {OUTPUT} ({size:.1f} MB)")
print(f"   อัพเดท: {out['updated']}")
print(f"{'='*40}")
