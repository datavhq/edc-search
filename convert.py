"""
EDC Parameter Config — Excel to JSON Converter v3
วิธีใช้: python convert.py
"""
import json, os, glob
import pandas as pd
from datetime import datetime

def safe(v):
    if v is None: return ""
    s = str(v).strip()
    if s in ("None","nan","NaT","NaN","","no parameter in VHQ","no paramater VHQ","Inisis Data"): return ""
    if s.endswith(".0"):
        try: return str(int(float(s)))
        except: pass
    return s.lstrip("'")

def yn(v):
    if v is None: return False
    s = str(v).strip().upper()
    return s in ("Y","YES","1","TRUE","X","✓","ENABLED","ENABLE")

def isnum(v):
    try: int(str(v).strip().replace('.0','')); return True
    except: return False

def save(bank, rows):
    updated = datetime.now().strftime("%Y-%m-%d %H:%M")
    out = {"bank":bank,"updated":updated,"total":len(rows),"records":rows}
    fname = f"data_{bank}.json"
    with open(fname,"w",encoding="utf-8") as f:
        json.dump(out,f,ensure_ascii=False,separators=(",",":"))
    size = os.path.getsize(fname)/1024/1024
    print(f"  💾 {fname} ({size:.1f} MB)")

def conn_type_kbank(tel, ip_gprs):
    if str(tel).upper().strip() != 'GPRS': return safe(tel)
    ip = safe(ip_gprs)
    if ip.startswith('172.29'): return 'AIS'
    if ip.startswith('172.30'): return 'DTAC'
    return safe(tel)

def limit_label(v, bank=''):
    s = safe(v).replace(',','').replace(' ','').replace('-','')
    mapping = {
        '9999999':'หลักหมื่น','99999999':'หลักแสน',
        '1999999':'หลักแสน','999999':'หลักแสน',
        '999999999':'หลักล้าน','9999999999':'หลักสิบล้าน',
        '200000000':'2 ล้าน','300000000':'3 ล้าน',
    }
    # check range like "1-999,999"
    if '999,999' in safe(v) or s == '999999': return 'หลักแสน'
    if '9,999,999' in safe(v) or s == '9999999': 
        if bank in ('KBANK','WOM'): return 'หลักล้าน'
        return 'หลักหมื่น'
    if '99,999,999' in safe(v) or s == '99999999':
        if bank in ('BAY','BBL','SCB'): return 'หลักแสน'
        return 'หลักแสน'
    return mapping.get(s, safe(v))

print("="*45)
print("EDC Search — Excel to JSON Converter v3")
print("="*45)

# ════ KBANK ════
files = glob.glob("*KBANK*.xlsx")
if files:
    print(f"\n[KBANK] {files[0]}")
    df = pd.read_excel(files[0], dtype=str, sheet_name="ParameterReport")
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("TID",""))
        if not tid or tid=="0": continue
        # hosts
        hosts = []
        host_cols = [(c,c) for c in df.columns if (c.startswith("TID_") or c.startswith("MID_"))]
        # group by name
        names = {}
        for c in df.columns:
            if c.startswith("TID_"):
                name = c[4:]
                t = safe(r.get(c,""))
                m = safe(r.get(f"MID_{name}",""))
                if t and isnum(t):
                    names[name] = {"tid":t,"mid":m if isnum(m) else ""}
        hosts = [{"name":k,"tid":v["tid"],"mid":v["mid"]} for k,v in names.items()]

        rows.append({
            "bank":"KBANK","tid":tid,
            "mid":safe(r.get("MID","")),
            "sn":safe(r.get("S/N EDC","")),
            "slip1":safe(r.get("Slip Line_1","")),
            "slip2":safe(r.get("Slip Line_2","")),
            "slip3":safe(r.get("Slip Line_3","")),
            "slip4":safe(r.get("Slip Line_4","")),
            "model":safe(r.get("Model","")),
            "sw":safe(r.get("SW Version","")),
            "settle":safe(r.get("Settle Mode","")),
            "time_settle":safe(r.get("Time settle","")),
            "limit":limit_label(r.get("จำกัดยอดรูด-swipe limit",""),'KBANK'),
            "linkpos":yn(r.get("ECR_Enabled","")),
            "erm":yn(r.get("ERM_Enabled","")),
            "ip":safe(r.get("IP Host Internet / GPRS #1","")),
            "conn_type":conn_type_kbank(r.get("Tel_system",""), r.get("IP Host Internet / GPRS #1","")),
            "apn":safe(r.get("APN SIM","")),
            "flow_wallet":safe(r.get("Flow_Type_Wallet","")),
            "keyin":yn(r.get("Function_KBANK_Key-in.","")),
            "tip":yn(r.get("Function_KBANK_Tip.","")),
            "preauth":yn(r.get("Function_KBANK_Preauth.","")),
            "offline":yn(r.get("Function_KBANK_OFFline.","")),
            "refund":yn(r.get("Function_KBANK_ReFund.","")),
            "multi":safe(r.get("Multi_mechant","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("KBANK", rows)

# ════ WOM ════
files = glob.glob("*WOM*.xlsx") or glob.glob("*Common_Template*.xlsx")
if files:
    print(f"\n[WOM] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("MAIN_TID",""))
        if not tid or tid=="0": continue
        # hosts
        names = {}
        for c in df.columns:
            if c.startswith("TID_"):
                name = c[4:]
                t = safe(r.get(c,""))
                m = safe(r.get(f"MID_{name}",""))
                if t and isnum(t):
                    names[name] = {"tid":t,"mid":m if isnum(m) else ""}
        hosts = [{"name":k,"tid":v["tid"],"mid":v["mid"]} for k,v in names.items()]

        settle_raw = safe(r.get("SETTLE_MODE",""))
        settle = "Auto Settlement" if settle_raw=="1" else "Force Settlement" if settle_raw=="2" else settle_raw

        fw_raw = safe(r.get("FLOW_TYPE_WALLET",""))
        flow_wallet = "B SCAN C" if fw_raw=="0" else "C SCAN B" if fw_raw=="1" else fw_raw

        conn = conn_type_kbank(r.get("TEL_SYSTEM",""), r.get("IP_HOST_INTERNET_GPRS_1",""))

        rows.append({
            "bank":"WOM","tid":tid,
            "mid":safe(r.get("MAIN_MID","")),
            "sn":safe(r.get("SERIAL_NO","")),
            "slip1":safe(r.get("SLIP_LINE_1","")),
            "slip2":safe(r.get("SLIP_LINE_2","")),
            "slip3":safe(r.get("SLIP_LINE_3","")),
            "slip4":safe(r.get("SLIP_LINE_4","")),
            "model":safe(r.get("MODEL","")),
            "sw":safe(r.get("SW_VERSION","")),
            "settle":settle,
            "time_settle":safe(r.get("TIME_SET","")),
            "limit":limit_label(r.get("LIMIT_AMOUNT",""),'WOM'),
            "linkpos":yn(r.get("FECR","")),
            "erm":yn(r.get("INIT_ERM","")),
            "ip":safe(r.get("IP_HOST_INTERNET_GPRS_1","")),
            "conn_type":conn,
            "apn":safe(r.get("APN_SIM","")),
            "flow_wallet":flow_wallet,
            "keyin":yn(r.get("FUNCTION_KBANK_KEYIN","")),
            "tip":yn(r.get("FUNCTION_KBANK_TIP","")),
            "preauth":yn(r.get("FUNCTION_KBANK_PRE_AUTH","")),
            "offline":yn(r.get("FUNCTION_KBANK_OFFLINE","")),
            "refund":yn(r.get("FUNCTION_KBANK_REFUND","")),
            "multi":safe(r.get("MULTI_MERCHANT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("WOM", rows)

# ════ BAY ════
files = glob.glob("*BAY*.xlsx")
if files:
    print(f"\n[BAY] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("TIDหลัก",""))
        if not tid or tid=="0": continue

        # connection type logic
        port_sec = safe(r.get("PORT ทำรายการ Secondary",""))
        ip_pri = safe(r.get("IP ทำรายการ Primary",""))
        port_pri = safe(r.get("PORT ทำรายการ Primary",""))
        if port_sec == "6007": conn = "AIS"
        elif port_sec == "6008": conn = "DTAC"
        elif port_pri == "5001" and ip_pri.startswith("172.30"): conn = "DTAC"
        elif port_pri == "5001" and ip_pri.startswith("10.0.24"): conn = "AIS"
        else: conn = "LAN"

        # hosts — group ACQUIRER sets
        hosts = []
        for i in range(1,10):
            for j in [''] + [str(x) for x in range(1,9)]:
                suffix = f"{i}{j}" if j else str(i)
                acq_col = f"ACQUIRER{suffix}"
                tid_col = f"TID_ACQUIRER{suffix}"
                mid_col = f"MID_ACQUIRER{suffix}"
                acq = safe(r.get(acq_col,""))
                t = safe(r.get(tid_col,""))
                m = safe(r.get(mid_col,""))
                if t and isnum(t):
                    hosts.append({"name":acq or acq_col,"tid":t,"mid":m if isnum(m) else ""})
        # ARKE/KBANK/KTC
        for grp in ['ARKE','KBANK','KTC']:
            for i in range(1,8):
                acq = safe(r.get(f"ACQUIRER{i}_{grp}",""))
                t = safe(r.get(f"TID{i}_{grp}",""))
                m = safe(r.get(f"MID{i}_{grp}",""))
                if t and isnum(t):
                    hosts.append({"name":acq or f"{grp}{i}","tid":t,"mid":m if isnum(m) else ""})

        rows.append({
            "bank":"BAY","tid":tid,
            "mid":safe(r.get("MID_ACQUIRER1","")),
            "sn":safe(r.get("serialNumber","")),
            "slip1":safe(r.get("SLIP1","")),
            "slip2":safe(r.get("SLIP2","")),
            "slip3":safe(r.get("SLIP3","")),
            "slip4":"",
            "model":safe(r.get("PROFILE","")),
            "sw":"",
            "settle":"",
            "time_settle":safe(r.get("AUTO SETTLEMENT TIME","")),
            "limit":limit_label(r.get("AMOUNT FORMAT",""),'BAY'),
            "linkpos":yn(r.get("LINK POS","")),
            "erm":"",
            "ip_pri":ip_pri,
            "ip_sec":safe(r.get("IP ทำรายการ Secondary","")),
            "conn_type":conn,
            "apn":"","flow_wallet":"",
            "keyin":yn(r.get("KEYIN","")),
            "tip":yn(r.get("ADJUST","")),
            "preauth":yn(r.get("PREAUTH","")),
            "offline":yn(r.get("OFFLINE","")),
            "refund":yn(r.get("REFUND","")),
            "multi":safe(r.get("MULTI-MERCHANT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("BAY", rows)

# ════ BBL ════
files = glob.glob("*BBL*.xlsx")
if files:
    print(f"\n[BBL] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("Terminal ID",""))
        if not tid or tid=="0": continue
        hosts = []
        hmap = {
            "BBL TID":"BBL","TID UPI":"UPI","TID-DCC":"DCC",
            "TID-UPI-DCI":"UPI-DCI","TID-BE-SMART":"BE-SMART",
            "TID-REDEMPTION":"REDEMPTION","TID.QR REF1 (TID)":"QR REF1",
            "TID_ALIPAY + WECHAT":"ALIPAY+WECHAT","TID BSS":"BSS",
        }
        mid_map = {
            "BBL TID":"BBL MID","TID UPI":"","TID-DCC":"",
            "TID-UPI-DCI":"","TID-BE-SMART":"MID-BE-SMART",
            "TID-REDEMPTION":"MID-REDEMPTION","TID.QR REF1 (TID)":"",
            "TID_ALIPAY + WECHAT":"","TID BSS":"MID BSS",
        }
        for tc, name in hmap.items():
            t = safe(r.get(tc,""))
            mc = mid_map.get(tc,"")
            m = safe(r.get(mc,"")) if mc else ""
            if t and isnum(t):
                hosts.append({"name":name,"tid":t,"mid":m if isnum(m) else ""})

        rows.append({
            "bank":"BBL","tid":tid,
            "mid":safe(r.get("BBL MID","")),
            "sn":safe(r.get("Serial No.","")),
            "slip1":safe(r.get("LINE 1","")),
            "slip2":safe(r.get("LINE 2","")),
            "slip3":safe(r.get("LINE 3","")),
            "slip4":safe(r.get("LINE 4","")),
            "model":safe(r.get("Model","")),
            "sw":"",
            "force_settle":safe(r.get("Force Settlement","")),
            "auto_settle":safe(r.get("Auto Settlement","")),
            "time_settle":"",
            "limit":limit_label(r.get("Transaction Limit",""),'BBL'),
            "linkpos":yn(r.get("LINKPOS","")),
            "erm":"",
            "ip":"","conn_type":"","apn":"","flow_wallet":"",
            "keyin":yn(r.get("Manual Key-in","")),
            "tip":yn(r.get("Adj. Tip","")),
            "preauth":yn(r.get("Preauth","")),
            "offline":yn(r.get("Offline Sale","")),
            "refund":yn(r.get("Refund","")),
            "multi":"",
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("BBL", rows)

# ════ SCB ════
SCB_HOST_NAMES = {
    1:"SCB",2:"UPI",3:"DCC",4:"IPP",5:"REDEEM",
    6:"ALIPAY",7:"WECHAT",8:"QR PROMPTPAY",9:"QRCS",
    10:"AMEX",11:"MIR",12:"TRUEMONEY"
}
files = glob.glob("*SCB*.xlsx")
if files:
    print(f"\n[SCB] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("HOST.1/TERMINAL_ID",""))
        if not tid or tid=="0": continue

        auto = safe(r.get("HOST.1/ENABLE_AUTO_SETTLEMENT",""))
        force = safe(r.get("HOST.1/ENABLE_FORCE_SETTLEMENT",""))
        if auto=="1" and force=="1": settle="AUTO+FORCE"
        elif auto=="1": settle="AUTO SETTLE"
        elif force=="1": settle="FORCE SETTLE"
        else: settle=""

        time_settle = ""
        if auto=="1": time_settle = safe(r.get("HOST.1/AUTO_SETTLE_TIME",""))
        elif force=="1": time_settle = safe(r.get("HOST.1/FORCE_SETTLE_TIME",""))

        gprs_ip = safe(r.get("CONNECTION.1/GPRS_PRIMARY_IP",""))
        if gprs_ip.startswith("172.29"): conn="AIS"
        elif gprs_ip.startswith("192.168"): conn="DTAC"
        else: conn="LAN"

        hosts = []
        for i in range(1,13):
            t = safe(r.get(f"HOST.{i}/TERMINAL_ID",""))
            m = safe(r.get(f"HOST.{i}/MERCHANT_ID",""))
            if t and isnum(t):
                name = SCB_HOST_NAMES.get(i, f"HOST.{i}")
                hosts.append({"name":name,"tid":t,"mid":m if isnum(m) else ""})

        rows.append({
            "bank":"SCB","tid":tid,
            "mid":safe(r.get("HOST.1/MERCHANT_ID","")),
            "sn":safe(r.get("serialNumber","")),
            "slip1":safe(r.get("PRINT_CONFIG.1/HEADER1","")),
            "slip2":safe(r.get("PRINT_CONFIG.1/HEADER2","")),
            "slip3":safe(r.get("PRINT_CONFIG.1/HEADER3","")),
            "slip4":"",
            "model":safe(r.get("modelName","")),
            "sw":"",
            "settle":settle,
            "time_settle":time_settle,
            "limit":limit_label(r.get("TERMINAL/MAX_TRANS_AMOUNT",""),'SCB'),
            "linkpos":yn(r.get("TERMINAL/ENABLE_ECR","")),
            "erm":"",
            "ip":safe(r.get("CONNECTION.1/PRIMARY_IP","")),
            "conn_type":conn,
            "apn":"","flow_wallet":"",
            "keyin":yn(r.get("TRANS_SWITCH.1/ENABLE_MANUAL_KEY_IN","")),
            "tip":yn(r.get("TERMINAL/ALLOW_TIP","")),
            "preauth":yn(r.get("TRANS_SWITCH.5/IS_SUPPORT_TRANS","")),
            "offline":yn(r.get("TRANS_SWITCH.4/IS_SUPPORT_TRANS","")),
            "refund":yn(r.get("TRANS_SWITCH.3/IS_SUPPORT_TRANS","")),
            "multi":safe(r.get("TERMINAL/ENABLE_MULTI_MERCHANT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("SCB", rows)

print("\n"+"="*45)
print("✅ เสร็จแล้ว! Upload ไฟล์เหล่านี้ขึ้น GitHub:")
for b in ["KBANK","WOM","BAY","BBL","SCB"]:
    f=f"data_{b}.json"
    if os.path.exists(f):
        print(f"  📄 {f} ({os.path.getsize(f)/1024/1024:.1f} MB)")
print("="*45)
