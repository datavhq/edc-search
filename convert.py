"""
EDC Parameter Config — Excel to JSON Converter v4
"""
import json, os, glob
import pandas as pd
from datetime import datetime

def safe(v):
    if v is None: return ""
    s = str(v).strip()
    if s in ("None","nan","NaT","NaN","","no parameter in VHQ","no paramater VHQ","Inisis Data","-","N/A"): return ""
    if s.endswith(".0"):
        try: return str(int(float(s)))
        except: pass
    return s.lstrip("'")

def yn(v):
    if v is None: return False
    s = str(v).strip().upper()
    return s in ("Y","YES","1","TRUE","X","✓","ENABLED","ENABLE")

def isnum(v):
    s = safe(v)
    if not s: return False
    try: int(s.replace(',','')); return True
    except: return False

def model_from_sn(sn, fallback=""):
    s = safe(sn)
    if s.upper().startswith("V1E"): return "X990 PCI5"
    if s.upper().startswith("V9E"): return "X990 PCI6"
    return safe(fallback) if fallback else s[:3] if s else ""

def limit_label(v, bank=''):
    s = safe(v).replace(',','').replace(' ','').replace('-','')
    if '999,999' in safe(v) or s in ('999999','99999999') and bank in ('KBANK','WOM'): return 'หลักแสน'
    if s == '99999999': return 'หลักแสน'
    if s == '999999999': return 'หลักล้าน'
    if s == '9999999': return 'หลักหมื่น'
    if s == '9999999999': return 'หลักสิบล้าน'
    if s == '200000000': return '2 ล้าน'
    if s == '300000000': return '3 ล้าน'
    return safe(v)

def conn_type_gprs(tel, ip):
    if str(tel).upper().strip() != 'GPRS': return safe(tel)
    if safe(ip).startswith('172.29'): return 'AIS'
    if safe(ip).startswith('172.30'): return 'DTAC'
    return safe(tel)

def save(bank, rows):
    updated = datetime.now().strftime("%Y-%m-%d %H:%M")
    out = {"bank":bank,"updated":updated,"total":len(rows),"records":rows}
    fname = f"data_{bank}.json"
    with open(fname,"w",encoding="utf-8") as f:
        json.dump(out,f,ensure_ascii=False,separators=(",",":"))
    size = os.path.getsize(fname)/1024/1024
    print(f"  💾 {fname} ({size:.1f} MB)")

print("="*48)
print("EDC Search — Excel to JSON Converter v4")
print("="*48)

# ══════════════════════════════════════════
# KBANK
# ══════════════════════════════════════════
files = glob.glob("*KBANK*.xlsx")
if files:
    print(f"\n[KBANK] {files[0]}")
    df = pd.read_excel(files[0], dtype=str, sheet_name="ParameterReport")
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("TID",""))
        if not tid or tid=="0": continue
        sn = safe(r.get("S/N EDC",""))
        model = model_from_sn(sn, r.get("Model",""))

        # hosts — ทุก TID_xxx / MID_xxx col + FULLPAYMENT_SCB
        hosts = []
        skip = {'TID','MID'}
        seen = set()
        for c in df.columns:
            if c.startswith("TID_") or c == "TID.FULLPAYMENT_SCB":
                name = c[4:] if c.startswith("TID_") else "FULLPAYMENT_SCB"
                if name in seen: continue
                seen.add(name)
                t = safe(r.get(c,""))
                mid_col = f"MID_{name}" if c.startswith("TID_") else "MID.FULLPAYMENT_SCB"
                m = safe(r.get(mid_col,""))
                if t and isnum(t):
                    hosts.append({"name":name,"tid":t,"mid":m if isnum(m) else ""})
        # KBANK main
        main_t = safe(r.get("TID",""))
        main_m = safe(r.get("MID",""))
        if main_t and isnum(main_t):
            hosts.insert(0,{"name":"KBANK","tid":main_t,"mid":main_m if isnum(main_m) else ""})

        rows.append({
            "bank":"KBANK","tid":tid,"mid":safe(r.get("MID","")),
            "sn":sn,"model":model,"sw":safe(r.get("SW Version","")),
            "slip1":safe(r.get("Slip Line_1","")),
            "slip2":safe(r.get("Slip Line_2","")),
            "slip3":safe(r.get("Slip Line_3","")),
            "slip4":safe(r.get("Slip Line_4","")),
            "settle":safe(r.get("Settle Mode","")),
            "time_settle":safe(r.get("Time settle","")),
            "limit":limit_label(r.get("จำกัดยอดรูด-swipe limit",""),"KBANK"),
            "linkpos":yn(r.get("ECR_Enabled","")),
            "erm":yn(r.get("ERM_Enabled","")),
            "ip":safe(r.get("IP Host Internet / GPRS #1","")),
            "conn_type":conn_type_gprs(r.get("Tel_system",""),r.get("IP Host Internet / GPRS #1","")),
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

# ══════════════════════════════════════════
# WOM
# ══════════════════════════════════════════
files = glob.glob("*WOM*.xlsx") or glob.glob("*Common_Template*.xlsx")
if files:
    print(f"\n[WOM] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("MAIN_TID",""))
        if not tid or tid=="0": continue
        sn = safe(r.get("SERIAL_NO",""))
        model = model_from_sn(sn, r.get("MODEL",""))

        hosts = []
        seen = set()
        for c in df.columns:
            if c.startswith("TID_"):
                name = c[4:]
                if name in seen: continue
                seen.add(name)
                t = safe(r.get(c,""))
                m = safe(r.get(f"MID_{name}",""))
                if t and isnum(t):
                    hosts.append({"name":name,"tid":t,"mid":m if isnum(m) else ""})

        settle_raw = safe(r.get("SETTLE_MODE",""))
        settle = "Auto Settlement" if settle_raw=="1" else "Force Settlement" if settle_raw=="2" else settle_raw
        fw_raw = safe(r.get("FLOW_TYPE_WALLET",""))
        flow_wallet = "B SCAN C" if fw_raw=="0" else "C SCAN B" if fw_raw=="1" else fw_raw
        conn = conn_type_gprs(r.get("TEL_SYSTEM",""), r.get("IP_HOST_INTERNET_GPRS_1",""))

        rows.append({
            "bank":"WOM","tid":tid,"mid":safe(r.get("MAIN_MID","")),
            "sn":sn,"model":model,"sw":safe(r.get("SW_VERSION","")),
            "slip1":safe(r.get("SLIP_LINE_1","")),
            "slip2":safe(r.get("SLIP_LINE_2","")),
            "slip3":safe(r.get("SLIP_LINE_3","")),
            "slip4":safe(r.get("SLIP_LINE_4","")),
            "settle":settle,"time_settle":safe(r.get("TIME_SET","")),
            "limit":limit_label(r.get("LIMIT_AMOUNT",""),"WOM"),
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

# ══════════════════════════════════════════
# BAY
# ══════════════════════════════════════════
files = glob.glob("*BAY*.xlsx")
if files:
    print(f"\n[BAY] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("TIDหลัก",""))
        if not tid or tid=="0": continue
        sn = safe(r.get("serialNumber",""))
        model = model_from_sn(sn)
        sw = safe(r.get("PROFILE",""))

        # connection type
        port_sec = safe(r.get("PORT ทำรายการ Secondary",""))
        ip_pri   = safe(r.get("IP ทำรายการ Primary",""))
        port_pri = safe(r.get("PORT ทำรายการ Primary",""))
        if port_sec=="6007": conn="AIS"
        elif port_sec=="6008": conn="DTAC"
        elif port_pri=="5001" and ip_pri.startswith("172.30"): conn="DTAC"
        elif port_pri=="5001" and ip_pri.startswith("10.0.24"): conn="AIS"
        else: conn="LAN"

        # hosts — ACQUIRER sets (1-9 x 1-6)
        hosts = []
        seen_tids = set()
        # main ACQUIRER groups
        for i in range(1,10):
            for j in [''] + [str(x) for x in range(1,7)]:
                suffix = f"{i}{j}" if j else str(i)
                acq = safe(r.get(f"ACQUIRER{suffix}",""))
                t   = safe(r.get(f"TID_ACQUIRER{suffix}",""))
                m   = safe(r.get(f"MID_ACQUIRER{suffix}",""))
                if t and isnum(t) and t not in seen_tids:
                    seen_tids.add(t)
                    hosts.append({"name":acq or f"ACQUIRER{suffix}","tid":t,"mid":m if isnum(m) else ""})
        # ARKE
        for i in range(1,8):
            acq = safe(r.get(f"ACQUIRER{i}_ARKE",""))
            t   = safe(r.get(f"TID{i}_ARKE",""))
            m   = safe(r.get(f"MID{i}_ARKE",""))
            if t and isnum(t) and t not in seen_tids:
                seen_tids.add(t)
                hosts.append({"name":acq or f"ARKE{i}","tid":t,"mid":m if isnum(m) else ""})
        # KBANK
        for i in range(1,8):
            acq = safe(r.get(f"ACQUIRER{i}_KBANK",""))
            t   = safe(r.get(f"TID{i}_KBANK",""))
            m   = safe(r.get(f"MID{i}_KBANK",""))
            if t and isnum(t) and t not in seen_tids:
                seen_tids.add(t)
                hosts.append({"name":acq or f"KBANK{i}","tid":t,"mid":m if isnum(m) else ""})
        # KTC
        for i in range(1,4):
            acq = safe(r.get(f"ACQUIRER{i}_KTC",""))
            t   = safe(r.get(f"TID{i}_KTC",""))
            m   = safe(r.get(f"MID{i}_KTC",""))
            if t and isnum(t) and t not in seen_tids:
                seen_tids.add(t)
                hosts.append({"name":acq or f"KTC{i}","tid":t,"mid":m if isnum(m) else ""})

        rows.append({
            "bank":"BAY","tid":tid,"mid":safe(r.get("MID_ACQUIRER1","")),
            "sn":sn,"model":model,"sw":sw,
            "slip1":safe(r.get("SLIP1","")),
            "slip2":safe(r.get("SLIP2","")),
            "slip3":safe(r.get("SLIP3","")),
            "slip4":"",
            "settle":"","time_settle":safe(r.get("AUTO SETTLEMENT TIME","")),
            "limit":limit_label(r.get("AMOUNT FORMAT",""),"BAY"),
            "linkpos":yn(r.get("LINK POS","")),
            "erm":"",
            "ip_pri":ip_pri,"ip_sec":safe(r.get("IP ทำรายการ Secondary","")),
            "conn_type":conn,"apn":"","flow_wallet":"",
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

# ══════════════════════════════════════════
# BBL
# ══════════════════════════════════════════
files = glob.glob("*BBL*.xlsx")
if files:
    print(f"\n[BBL] {files[0]}")
    df = pd.read_excel(files[0], dtype=str)
    rows = []
    for _, r in df.iterrows():
        tid = safe(r.get("Terminal ID",""))
        if not tid or tid=="0": continue
        sn = safe(r.get("Serial No.",""))
        model = model_from_sn(sn, r.get("Model",""))

        hosts = []
        hmap = [
            ("BBL","BBL TID","BBL MID"),
            ("UPI","TID UPI",""),
            ("DCC","TID-DCC",""),
            ("UPI-DCI","TID-UPI-DCI",""),
            ("BE-SMART","TID-BE-SMART","MID-BE-SMART"),
            ("REDEMPTION","TID-REDEMPTION","MID-REDEMPTION"),
            ("QR REF1","TID.QR REF1 (TID)",""),
            ("ALIPAY+WECHAT","TID_ALIPAY + WECHAT",""),
            ("BSS","TID BSS","MID BSS"),
        ]
        for name,tc,mc in hmap:
            t = safe(r.get(tc,""))
            m = safe(r.get(mc,"")) if mc else ""
            if t and isnum(t):
                hosts.append({"name":name,"tid":t,"mid":m if isnum(m) else ""})

        rows.append({
            "bank":"BBL","tid":tid,"mid":safe(r.get("BBL MID","")),
            "sn":sn,"model":model,"sw":"",
            "slip1":safe(r.get("LINE 1","")),
            "slip2":safe(r.get("LINE 2","")),
            "slip3":safe(r.get("LINE 3","")),
            "slip4":safe(r.get("LINE 4","")),
            "force_settle":safe(r.get("Force Settlement","")),
            "auto_settle":safe(r.get("Auto Settlement","")),
            "time_settle":"",
            "limit":limit_label(r.get("Transaction Limit",""),"BBL"),
            "linkpos":yn(r.get("LINKPOS","")),
            "erm":"","ip":"","conn_type":"","apn":"","flow_wallet":"",
            "keyin":yn(r.get("Manual Key-in","")),
            "tip":yn(r.get("Adj. Tip","")),
            "preauth":yn(r.get("Preauth","")),
            "offline":yn(r.get("Offline Sale","")),
            "refund":yn(r.get("Refund","")),
            "biller_id":safe(r.get("BILLER ID","")),
            "biller_name":safe(r.get("BILLER NAME","")),
            "multi":"",
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("BBL", rows)

# ══════════════════════════════════════════
# SCB
# ══════════════════════════════════════════
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
        tid = safe(r.get("deviceId","")) or safe(r.get("HOST.1/TERMINAL_ID",""))
        if not tid or tid=="0": continue
        sn = safe(r.get("serialNumber",""))
        model = model_from_sn(sn, r.get("modelName",""))

        # settle
        auto  = safe(r.get("HOST.1/ENABLE_AUTO_SETTLEMENT",""))
        force = safe(r.get("HOST.1/ENABLE_FORCE_SETTLEMENT",""))
        if auto=="1" and force=="1":   settle="AUTO+FORCE"
        elif auto=="1":                settle="AUTO SETTLE"
        elif force=="1":               settle="FORCE SETTLE"
        else:                          settle=""
        t_auto  = safe(r.get("HOST.1/AUTO_SETTLE_TIME",""))
        t_force = safe(r.get("HOST.1/FORCE_SETTLE_TIME",""))
        if auto=="1" and force=="1":   time_settle=f"{t_auto}+{t_force}" if t_auto and t_force else t_auto or t_force
        elif auto=="1":                time_settle=t_auto
        elif force=="1":               time_settle=t_force
        else:                          time_settle=""

        # conn
        gprs_ip = safe(r.get("CONNECTION.1/GPRS_PRIMARY_IP",""))
        if gprs_ip.startswith("172.29"):   conn="AIS"
        elif gprs_ip.startswith("192.168"): conn="DTAC"
        else:                               conn="LAN"

        hosts = []
        for i in range(1,13):
            t = safe(r.get(f"HOST.{i}/TERMINAL_ID",""))
            m = safe(r.get(f"HOST.{i}/MERCHANT_ID",""))
            if t and isnum(t):
                name = SCB_HOST_NAMES.get(i,f"HOST.{i}")
                hosts.append({"name":name,"tid":t,"mid":m if isnum(m) else ""})

        rows.append({
            "bank":"SCB","tid":tid,"mid":safe(r.get("HOST.1/MERCHANT_ID","")),
            "sn":sn,"model":model,"sw":"",
            "slip1":safe(r.get("PRINT_CONFIG.1/HEADER1","")),
            "slip2":safe(r.get("PRINT_CONFIG.1/HEADER2","")),
            "slip3":safe(r.get("PRINT_CONFIG.1/HEADER3","")),
            "slip4":"",
            "settle":settle,"time_settle":time_settle,
            "limit":limit_label(r.get("TERMINAL/MAX_TRANS_AMOUNT",""),"SCB"),
            "linkpos":yn(r.get("TERMINAL/ENABLE_ECR","")),
            "erm":"",
            "ip":safe(r.get("CONNECTION.1/PRIMARY_IP","")),
            "conn_type":conn,"apn":"","flow_wallet":"",
            "keyin":yn(r.get("TRANS_SWITCH.1/ENABLE_MANUAL_KEY_IN","")),
            "tip":yn(r.get("TERMINAL/ALLOW_TIP","")),
            "preauth":yn(r.get("TRANS_SWITCH.5/IS_SUPPORT_TRANS","")),
            "offline":yn(r.get("TRANS_SWITCH.4/IS_SUPPORT_TRANS","")),
            "refund":yn(r.get("TRANS_SWITCH.3/IS_SUPPORT_TRANS","")),
            "biller_name":safe(r.get("HOST.8/QR_MERCHANT_NAME","")),
            "biller_id_ref1":safe(r.get("QR_CONFIG.1/QR_BILLERID_ONLY_REF1","")),
            "biller_id_ref2":safe(r.get("QR_CONFIG.1/QR_BILLERID_WITH_REF2","")),
            "multi":safe(r.get("TERMINAL/ENABLE_MULTI_MERCHANT","")),
            "hosts":hosts,
        })
    print(f"  ✓ {len(rows):,} records")
    save("SCB", rows)

print("\n"+"="*48)
print("✅ เสร็จแล้ว!")
for b in ["KBANK","WOM","BAY","BBL","SCB"]:
    f=f"data_{b}.json"
    if os.path.exists(f):
        print(f"  📄 {f} ({os.path.getsize(f)/1024/1024:.1f} MB)")
print("="*48)
