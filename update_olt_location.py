import os
import glob
import re
import pandas as pd
import json
from datetime import datetime

# Define base directory and target directories
BASE_DIR = '/Users/bbae/GPTCodex'
TARGET_DIRS = [
    '/Users/bbae/OLT-Location',
    '/Users/bbae/GPTCodex/OLT-Location',
    '/Users/bbae/GPTCodex/OLT-Location/GIT-LOCATION'
]

def extract_date_from_filename(filename):
    if not filename:
        return None
    fn = os.path.basename(filename)
    if fn.startswith('~$'):
        return None
    # 1. Match YYYYMMDD (e.g. Device-20260829.xlsx -> 2026-08-29, OLT_EDFA_XGPON_Master_Inventory_Report_20260901.xlsx)
    m = re.search(r'(20\d{2})(\d{2})(\d{2})', fn)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except:
            pass
    # 1.1 Match YYYYMM (e.g. QRun_Corporate_202609.xlsx -> 2026-09-01)
    m = re.search(r'(20\d{2})(\d{2})(?!\d)', fn)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), 1)
        except:
            pass
    # 2. Match DD-MM-YYYY or DD_MM_YYYY (e.g. Report_23-08-2026.csv)
    m = re.search(r'(\d{2})[-_](\d{2})[-_](20\d{2})', fn)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except:
            pass
    # 3. Match DDMMYY or DDMMYYYY (e.g. MaxSpeed270826.xlsx -> 27/08/26, PonPort-280826.xlsx, MANE010926.csv)
    m = re.search(r'(\d{2})(\d{2})(\d{2,4})', fn)
    if m:
        d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime(y, mth, d)
        except:
            pass
    try:
        return datetime.fromtimestamp(os.path.getmtime(filename))
    except:
        return None

def find_latest_file(patterns):
    matched = []
    for p in patterns:
        matched.extend(glob.glob(p))
    matched = [f for f in matched if not os.path.basename(f).startswith('~$') and os.path.isfile(f) and os.path.getsize(f) > 0]
    def sort_key(f):
        dt = extract_date_from_filename(f)
        if dt:
            return dt
        return datetime.fromtimestamp(os.path.getmtime(f))
    if not matched:
        return None
    matched.sort(key=sort_key, reverse=True)
    return matched[0]

# Models that are fixed box / 1-slot built-in modules (never show Empty Slot)
FIXED_BOX_MODELS = {
    'C610', 'AN6001-G16', 'AN6001-GF16', 'MA5801-FL16', 'MA5801',
    '7362 ISAM DF-8G', 'LS-DF-CFXR-F', 'GL5610-08P(V2)', 'GL5610-04P(V3)',
    'ISCOM5504', 'ISCOM5508', 'V5804', 'V5808'
}

# Standard Chassis PON Service Slot Definitions for Modular Models
CHASSIS_SERVICE_SLOTS = {
    # 2-Slot Modular Models
    'MA5800-X2': ['0-1', '0-2'],
    'AN6000-2': ['1-1', '1-2'],
    'MA5608T': ['0-0', '0-1'],
    'AN5516-04': ['1-1', '1-2'],
    'C320': ['1-1', '1-2'],
    'C620': ['1-1', '1-2'],
    
    # High-density Modular Models
    'MA5800-X17': ['0-1', '0-2', '0-3', '0-4', '0-5', '0-6', '0-7', '0-8', '0-11', '0-12', '0-13', '0-14', '0-15', '0-16', '0-17', '0-18'],
    'C300': ['1-2', '1-3', '1-4', '1-5', '1-6', '1-7', '1-8', '1-9', '1-12', '1-13', '1-14', '1-15', '1-16', '1-17', '1-18', '1-19'],
    'MA5603T': ['0-0', '0-1', '0-2', '0-3', '0-4', '0-5'],
    'AN6000-17': [f'1-{i}' for i in range(1, 18) if i not in [9, 10]],
    'C650': [f'1-{i}' for i in range(1, 8)],
}

def is_pon_service_card(board_type, role):
    bt = str(board_type).upper().strip()
    r = str(role).upper().strip()
    
    # Exclude non-service roles
    if any(x in r for x in ['CTRL', 'POWER', 'PWR', 'FAN', 'UPLINK']):
        return False
        
    # Explicit non-service board models/types
    non_service_patterns = [
        # Fan
        'FAN', 'PIDD', 'FCSDA', 'FUMO',
        # Power
        'PWR', 'POWER', 'PDC', 'MPWC', 'MPWD', 'PRTE', 'PRTG', 'PRWG', 'PRWH', 'PRSF', 'PDC300', 'PUMD',
        'PILA', 'PISC', 'PISA', 'PWRD',
        # Control / Main Processing / Uplink
        'MCUD', 'SCUN', 'SCUL', 'MPLB', 'MPSC', 'MPSG', 'MPSA', 'SMXA', 'HSUD', 'HU6P', 'HSUB', 'HU2P', 
        'HMPA', 'HMPB', 'SCTMB', 'SPUFM', 'SPUF', 'DFMB-C', 'LMNT-A', 'NXHC', 'MPCB', 'CIUA',
        'GICF', 'CITD', 'CCKRC', 'CDAC', 'HSCV', 'HSCP',
        # VDSL / Voice / POTS
        'VDPE', 'SHLM', 'ASPB', 'ASP', 'VDP',
        # Fiberhome internal power / non-service numeric types
        '611', '640', '639', '389', '589'
    ]
    if any(p in bt for p in non_service_patterns):
        return False
        
    # PON card identifiers
    pon_patterns = [
        # Huawei GPON/XGS
        'GPHF', 'GPFD', 'GPBH', 'GPBD', 'CSHF', 'FLHF', 'CGHF', 'GPON', 'XGS',
        # ZTE PON (GTGH, GTGO, GTGK, GFGL, GFGO, GFGH, HFTL)
        'GTG', 'GFG', 'HFTL',
        # Fiberhome PON (GC8B, GPOA, GPOP, GFOA, GVHH, GOOB, etc.)
        'GC8B', 'GPOA', 'GPOP', 'GFOA', 'GVHH', 'GOOB',
        # Nokia PON
        'NGLT', 'FGNT', 'FGLT',
        # Generic
        'PON'
    ]
    return any(p in bt for p in pon_patterns)

def normalize_uplink_sitename(s):
    s = str(s).strip()
    if not s or s.lower() in ['nan', 'none', '-']:
        return ''
    
    # 1. If explicit site in parentheses: e.g. CPE-ATG0016_(ATG6703) -> CPE-ATG6703
    m_par = re.search(r'\(([A-Z0-9_-]+)\)', s)
    if m_par:
        inner = m_par.group(1).strip()
        m_inner = re.search(r'([A-Z]{3}\d{4})', inner)
        if m_inner:
            m_pre = re.match(r'^(PN\d*|DN\d*|CPE|AN\d*|RN\d*|AGN\d*)', s, re.I)
            if m_pre:
                return f'{m_pre.group(1).upper()}-{m_inner.group(1).upper()}'
            return m_inner.group(1).upper()
        return inner

    # 2. Check if already format like PN1-PCT1291, PN2-PCT1291, CPE-ANC1602
    m_std = re.match(r'^(PN\d*|DN\d*|CPE|AN\d*|RN\d*|AGN\d*)-([A-Z]{3}\d{4})$', s, re.I)
    if m_std:
        return f'{m_std.group(1).upper()}-{m_std.group(2).upper()}'

    # 3. If contains underscore, e.g. PN-ANC1246-1_ACRACR0401M, DN-CBR6043-1_BLMCBI6000M, PCTPCT0400M_PCT1291
    if '_' in s:
        parts = s.split('_')
        left = parts[0].strip()
        right = parts[-1].strip()
        
        # Check right side first (e.g. PCT1291 in PCTPCT0400M_PCT1291)
        m_right = re.match(r'^([A-Z]{3}\d{4})$', right)
        if m_right:
            site_code = m_right.group(1).upper()
            m_pre = re.match(r'^(PN\d*|DN\d*|CPE|AN\d*|RN\d*|AGN\d*)', left, re.I)
            if m_pre:
                return f'{m_pre.group(1).upper()}-{site_code}'
            return site_code
            
        # Left side: e.g. PN-ANC1246-1, DN-CBR6043-1, AN-BGC04-1
        m_left = re.match(r'^([A-Z0-9]+)-([A-Z0-9]+)-?(\d+)?$', left, re.I)
        if m_left:
            prefix = m_left.group(1).upper()
            site = m_left.group(2).upper()
            idx = m_left.group(3)
            if idx and prefix in ['PN', 'DN', 'AN', 'RN', 'AGN']:
                return f'{prefix}{idx}-{site}'
            elif idx:
                return f'{prefix}-{site}-{idx}'
            return f'{prefix}-{site}'
        return left

    # 4. Standard dash pattern with trailing index: e.g. PN-ANC1246-1 -> PN1-ANC1246
    m_node = re.match(r'^(PN|DN|CPE|AN|RN|AGN)-([A-Z]{3}\d{4})-?(\d+)?$', s, re.I)
    if m_node:
        prefix = m_node.group(1).upper()
        site = m_node.group(2).upper()
        idx = m_node.group(3)
        if idx and prefix in ['PN', 'DN', 'AN', 'RN', 'AGN']:
            return f'{prefix}{idx}-{site}'
        return f'{prefix}-{site}'

    return s

def build_chassis_service_slots(model, actual_boards):
    # Fixed box / 1-slot built-in models: never add empty slot
    if any(m in str(model) for m in FIXED_BOX_MODELS):
        return [b for b in actual_boards if b.get('type') != 'Empty Slot']
        
    all_chassis_slots = CHASSIS_SERVICE_SLOTS.get(model, None)
    if not all_chassis_slots:
        return actual_boards
        
    actual_by_slot = {b['slot']: b for b in actual_boards}
    final_boards = []
    
    for slot in all_chassis_slots:
        if slot in actual_by_slot:
            final_boards.append(actual_by_slot[slot])
        else:
            final_boards.append({
                "slot": slot,
                "type": "Empty Slot",
                "role": "Service Slot (Available)",
                "status": "Empty / Standby"
            })
            
    # Include any extra unexpected slots if found
    for b in actual_boards:
        if b['slot'] not in all_chassis_slots:
            final_boards.append(b)
            
    return final_boards

def get_prop(name, prop_map, default=None):
    if not name or not prop_map:
        return default
    if name in prop_map:
        return prop_map[name]
    if 'GO' in name:
        alt = re.sub(r'GO([0-9])', r'G0\1', name)
        if alt in prop_map:
            return prop_map[alt]
    elif 'G0' in name:
        alt = re.sub(r'G0([0-9])', r'GO\1', name)
        if alt in prop_map:
            return prop_map[alt]
    return default

def main():
    # 1. Load historical EDFA models from existing JSON files
    historical_edfa = {}
    for tdir in TARGET_DIRS:
        jpath = os.path.join(tdir, 'olt_location_data.json')
        if os.path.exists(jpath):
            try:
                with open(jpath, 'r', encoding='utf-8') as f:
                    old_data = json.load(f)
                    for k, v in old_data.items():
                        if (v.get('edfa_model') or v.get('edfa_vendor')) and k not in historical_edfa:
                            historical_edfa[k] = {
                                'edfa_model': v.get('edfa_model', ''),
                                'edfa_vendor': v.get('edfa_vendor', ''),
                                'has_fwdm': v.get('has_fwdm', False),
                                'fwdm_clli': v.get('fwdm_clli', '')
                            }
            except Exception as e:
                print(f"Note loading historical JSON from {jpath}: {e}")
    print(f"Loaded {len(historical_edfa)} historical EDFA records for fallback.")

    # 2. Identify latest files dynamically
    loc_file = find_latest_file([
        os.path.join(BASE_DIR, 'inventory_port_olt_*.xlsx'),
        os.path.join(BASE_DIR, 'reports', 'inventory_port_olt_*.xlsx'),
        os.path.join(BASE_DIR, 'OLT-Location', 'inventory_port_olt_*.xlsx'),
        os.path.join(BASE_DIR, 'data', 'misc', 'OLT_Location*.xlsx')
    ])

    device_file = find_latest_file([
        os.path.join(BASE_DIR, 'Device-*.xlsx'),
        os.path.join(BASE_DIR, 'data', 'device', 'Device-*.xlsx')
    ])

    board_file = find_latest_file([
        os.path.join(BASE_DIR, 'Board-*.xlsx'),
        os.path.join(BASE_DIR, 'cardolt', 'Board-*.xlsx')
    ])

    speed_file = find_latest_file([
        os.path.join(BASE_DIR, 'OLT-Uplink-MaxSpeed*.xlsx'),
        os.path.join(BASE_DIR, 'oltmaxspeed', 'OLT-Uplink-MaxSpeed*.xlsx')
    ])

    corp_file = find_latest_file([
        os.path.join(BASE_DIR, 'QRun_Corporate_*.xlsx'),
        os.path.join(BASE_DIR, 'data', 'onu', 'QRun_Corporate_*.xlsx'),
        os.path.join(BASE_DIR, 'pon report', 'QRun_Corporate_*.xlsx')
    ])

    pon_file = find_latest_file([
        os.path.join(BASE_DIR, 'PonPort-*.xlsx'),
        os.path.join(BASE_DIR, 'reports', 'PonPort-*.xlsx')
    ])

    # Find latest OLT_EDFA_XGPON_Master_Inventory_Report (prioritize root first per Rule #6)
    edfa_file = find_latest_file([
        os.path.join(BASE_DIR, 'OLT_EDFA_XGPON_Master_Inventory_Report*.xlsx'),
        os.path.join(BASE_DIR, 'reports', 'OLT_EDFA_XGPON_Master_Inventory_Report*.xlsx')
    ])

    # Find latest TPP Consolidate Audit Report (verified outdoor cabinets from photos)
    tpp_audit_file = find_latest_file([
        os.path.join(BASE_DIR, 'TPP_Consolidate_OLT_EDFA_Audit_Report*.xlsx'),
        os.path.join(BASE_DIR, 'reports', 'TPP_Consolidate_OLT_EDFA_Audit_Report*.xlsx')
    ])

    # Find latest MANE file
    mane_file = find_latest_file([
        os.path.join(BASE_DIR, 'MANE*.csv'),
        os.path.join(BASE_DIR, 'oltmaxspeed', 'MANE*.csv')
    ])

    if not loc_file or not device_file or not board_file or not speed_file or not corp_file or not edfa_file:
        print("Error: Missing required files.")
        print(f"  loc_file:    {loc_file}")
        print(f"  device_file: {device_file}")
        print(f"  board_file:  {board_file}")
        print(f"  speed_file:  {speed_file}")
        print(f"  corp_file:   {corp_file}")
        print(f"  edfa_file:   {edfa_file}")
        return

    print(f"Using location file:    {loc_file}")
    print(f"Using device file:      {device_file}")
    print(f"Using board file:       {board_file}")
    print(f"Using speed file:       {speed_file}")
    print(f"Using corp file:        {corp_file}")
    print(f"Using PonPort file:     {pon_file}")
    print(f"Using EDFA Master file: {edfa_file}")
    print(f"Using TPP Audit file:   {tpp_audit_file}")
    print(f"Using MANE file:        {mane_file}")

    # 3. Load Device Master to map OLT name -> Vendor, Model, Type & IP (NMS Ground Truth)
    print("Loading Device master...")
    df_dev = pd.read_excel(device_file)
    device_models = {}
    device_vendors = {}
    device_types = {}
    device_original_names = {}
    device_ips = {}
    device_master_names = set()
    for _, row in df_dev.iterrows():
        raw_name = str(row.get('Device Name', '')).strip().upper()
        if not raw_name or raw_name == 'NAN':
            continue
        vendor = str(row.get('Vendor', '')).strip().upper()
        model = str(row.get('Model', '')).strip()
        typ = str(row.get('OLT Type', '')).strip().lower()
        dev_ip = str(row.get('Device IP', '')).strip()

        device_master_names.add(raw_name)
        device_models[raw_name] = model
        if vendor and vendor != 'NAN':
            device_vendors[raw_name] = vendor
        device_types[raw_name] = typ
        device_original_names[raw_name] = raw_name
        if dev_ip and dev_ip.lower() != 'nan':
            device_ips[raw_name] = dev_ip

    # 3.2. Load Corporate Links to count corp connections per OLT
    print("Loading Corporate links...")
    df_corp = pd.read_excel(corp_file)
    corp_counts = df_corp['OLT_NAME'].value_counts().to_dict()
    corp_map = {}
    for name, cnt in corp_counts.items():
        clean_name = str(name).strip().upper()
        corp_map[clean_name] = int(cnt)

    # 3.4. Load EDFA Model, Vendor and FWDM from OLT_EDFA_XGPON_Master_Inventory_Report
    print("Loading EDFA Model, Vendor and FWDM from OLT_EDFA_XGPON_Master_Inventory_Report...")
    edfa_model_map = {}
    edfa_vendor_map = {}
    edfa_count_map = {}
    fwdm_map = {}
    fwdm_clli_map = {}
    edfa_audited_olts = set()
    try:
        df_edfa = pd.read_excel(edfa_file, sheet_name=0)
        model_col = next((c for c in ['Paired EDFA Model(s)', 'EDFA Model', 'MODEL'] if c in df_edfa.columns), None)
        vendor_col = next((c for c in ['Paired EDFA Vendor(s)', 'EDFA Vendor', 'VENDOR'] if c in df_edfa.columns), None)
        count_col = next((c for c in ['Paired EDFA Count', 'EDFA Count'] if c in df_edfa.columns), None)
        fwdm_col = next((c for c in ['Has FWDM (MANE)', 'FWDM', 'Has FWDM'] if c in df_edfa.columns), None)
        fwdm_clli_col = next((c for c in ['FWDM CLLI Code', 'FWDM CLLI'] if c in df_edfa.columns), None)
        
        for _, row in df_edfa.iterrows():
            name = str(row.get('OLT Name', '')).strip().upper()
            name_std = re.sub(r'GO([0-9])', r'G0\1', name)
            if not name_std:
                continue
            edfa_audited_olts.add(name_std)
            edfa_audited_olts.add(name)
            
            if model_col:
                m_val = str(row.get(model_col, '')).strip()
                if m_val and m_val.lower() != 'nan':
                    edfa_model_map[name_std] = m_val
                    edfa_model_map[name] = m_val
                    
            if vendor_col:
                v_val = str(row.get(vendor_col, '')).strip()
                if v_val and v_val.lower() != 'nan':
                    edfa_vendor_map[name_std] = v_val
                    edfa_vendor_map[name] = v_val
                    
            if count_col:
                c_val = row.get(count_col)
                if pd.notna(c_val):
                    try:
                        edfa_count_map[name_std] = int(c_val)
                        edfa_count_map[name] = int(c_val)
                    except:
                        pass

            if fwdm_col:
                f_val = str(row.get(fwdm_col, '')).strip().upper()
                if f_val in ['YES', 'Y', 'TRUE', '1']:
                    fwdm_map[name_std] = True
                    fwdm_map[name] = True
            
            if fwdm_clli_col:
                fc_val = str(row.get(fwdm_clli_col, '')).strip()
                if fc_val and fc_val.lower() != 'nan':
                    fwdm_clli_map[name_std] = fc_val
                    fwdm_clli_map[name] = fc_val

        print(f"Successfully loaded {len(edfa_model_map)} EDFA Model mappings and {len(fwdm_map)} FWDM sites from master report.")
    except Exception as e:
        print(f"Error loading EDFA Master Report: {e}")

    # 3.4.2 Load verified Outdoor OLTs from TPP Consolidate Audit Report (and fallback set)
    print("Loading verified Outdoor OLTs from TPP Consolidate Audit...")
    tpp_outdoor_olts = set()
    if tpp_audit_file and os.path.exists(tpp_audit_file):
        try:
            df_tpp = pd.read_excel(tpp_audit_file, sheet_name=0)
            site_col = next((c for c in ['Site_Name_Consolidate', 'Consolidate_Site_Name', 'Site_Name'] if c in df_tpp.columns), None)
            olt_col = next((c for c in ['OLT_Name_Found', 'Normalized_OLT_Name', 'OLT_Name'] if c in df_tpp.columns), None)
            if site_col and olt_col:
                actual_tpp = df_tpp[~df_tpp[site_col].astype(str).str.contains('EXAMPLE|โกดัง', case=False, na=False)]
                for olt_name in actual_tpp[olt_col].dropna():
                    clean_olt = str(olt_name).strip().upper()
                    std_olt = re.sub(r'GO([0-9])', r'G0\1', clean_olt)
                    tpp_outdoor_olts.add(clean_olt)
                    tpp_outdoor_olts.add(std_olt)
            print(f"Successfully loaded {len(tpp_outdoor_olts)} verified Outdoor OLT identifiers from TPP audit file ({os.path.basename(tpp_audit_file)}).")
        except Exception as e:
            print(f"Error loading TPP Audit file: {e}")

    # Fallback to report/tpp_consolidate_outdoor_olts_293.txt if exists
    tpp_txt_path = os.path.join(BASE_DIR, 'report', 'tpp_consolidate_outdoor_olts_293.txt')
    if os.path.exists(tpp_txt_path):
        try:
            with open(tpp_txt_path, 'r', encoding='utf-8') as f:
                for line in f:
                    clean_olt = line.strip().upper()
                    if clean_olt:
                        std_olt = re.sub(r'GO([0-9])', r'G0\1', clean_olt)
                        tpp_outdoor_olts.add(clean_olt)
                        tpp_outdoor_olts.add(std_olt)
        except Exception as e:
            pass

    # 3.4.5 Load MANE put inservice dates and RCU_Uplink
    print("Loading MANE put inservice dates and uplink sites...")
    inservice_map = {}
    mane_uplink_sites = {}
    if mane_file:
        try:
            df_mane = pd.read_csv(mane_file, sep='|', dtype=str)
            for _, row in df_mane.iterrows():
                clli = str(row.get('CLLI_Code', '')).strip().upper()
                if not clli or clli == 'NAN':
                    continue
                status = str(row.get('Status', '')).strip()
                if status.lower() == 'put inservice':
                    date_val = str(row.get('ActionDate', '')).strip()
                    if date_val and date_val.lower() != 'nan':
                        inservice_map[clli] = date_val
                rcu = str(row.get('RCU_Uplink', '')).strip()
                if rcu and rcu.lower() != 'nan':
                    mane_uplink_sites[clli] = rcu
        except Exception as e:
            print(f"Error reading MANE file: {e}")

    # 3.5. Load OLT Uplink Speed, ONU Count and PON Port configuration
    print("Loading OLT Uplink speed, sites, and ports...")
    df_speed = pd.read_excel(speed_file)
    speed_map = {}
    speed_ip_map = {}
    speed_uplink_sites = {}
    onu_count_map = {}
    ports_use_map = {}
    ports_total_map = {}
    slots_map = {}
    for _, row in df_speed.iterrows():
        dev_name = str(row.get('Device', '')).strip().upper()
        if not dev_name or dev_name == 'NAN':
            continue
        speed_descr = str(row.get('MaxSpeedDescr', '')).strip()
        onu_cnt = int(row.get('ONU Count', 0)) if not pd.isna(row.get('ONU Count')) else 0
        p_used = int(row.get('PON Port Used', 0)) if not pd.isna(row.get('PON Port Used')) else 0
        p_total = int(row.get('PON Port Count', 0)) if not pd.isna(row.get('PON Port Count')) else 0
        slot_cnt = int(row.get('PON Slot Count', 0)) if not pd.isna(row.get('PON Slot Count')) else 0
        ip_val = str(row.get('IP', '')).strip()
        dst_site = str(row.get('DstSite', '')).strip()

        # Normalize speed
        norm_speed = "-"
        if "10" in speed_descr:
            norm_speed = "10G"
        elif "1" in speed_descr:
            norm_speed = "1G"
        elif speed_descr and speed_descr.lower() != 'nan':
            norm_speed = speed_descr
        speed_map[dev_name] = norm_speed
        if ip_val and ip_val.lower() != 'nan' and dev_name not in speed_ip_map:
            speed_ip_map[dev_name] = ip_val
        if dst_site and dst_site.lower() != 'nan':
            if dev_name not in speed_uplink_sites:
                speed_uplink_sites[dev_name] = []
            if dst_site not in speed_uplink_sites[dev_name]:
                speed_uplink_sites[dev_name].append(dst_site)
        onu_count_map[dev_name] = onu_cnt
        ports_use_map[dev_name] = p_used
        ports_total_map[dev_name] = p_total
        slots_map[dev_name] = slot_cnt

    # 3.6. Load PonPort to count plugged PON SFPs per OLT
    sfp_map = {}
    if pon_file:
        print("Loading PonPort for plugged SFP statistics...")
        try:
            df_pon = pd.read_excel(pon_file, usecols=['Device Name', 'ModuleClass', 'VendorPN', 'Model', 'Vendor'])
            df_pon['OLT_RAW'] = df_pon['Device Name'].astype(str).str.strip().str.upper()

            def check_plugged(row):
                mc = str(row['ModuleClass']).strip().lower()
                pn = str(row['VendorPN']).strip().lower()
                v = str(row['Vendor']).strip().lower()
                m = str(row['Model']).strip().lower()
                if mc in ['', 'nan', 'none', 'invalid-invalid', 'absent']:
                    if pn in ['', 'nan', 'none', 'absent', '-']:
                        return False
                if 'ma5801' in m:
                    if mc == 'gponandxgspon-invalid':
                        return True
                    if mc == 'invalid-invalid':
                        return False
                if 'zte' in v:
                    if mc in ['', 'nan', 'absent'] and pn in ['', 'nan', 'absent', '-']:
                        return False
                if mc not in ['', 'nan', 'none', 'invalid-invalid', 'absent'] or pn not in ['', 'nan', 'none', 'absent', '-']:
                    return True
                return False

            df_pon['is_plugged'] = df_pon.apply(check_plugged, axis=1)
            df_pon['is_combo'] = df_pon['is_plugged'] & (
                df_pon['ModuleClass'].astype(str).str.lower().str.contains('xgs|combo|xg-pon|10g|gponandxgspon', regex=True) |
                df_pon['VendorPN'].astype(str).str.upper().str.contains('Z0404|Z0407', regex=True)
            )

            agg_sfp = df_pon.groupby('OLT_RAW').agg(
                plugged=('is_plugged', 'sum'),
                combo=('is_combo', 'sum')
            ).reset_index()
            agg_sfp['gpon'] = agg_sfp['plugged'] - agg_sfp['combo']

            for _, r in agg_sfp.iterrows():
                sfp_map[r['OLT_RAW']] = {
                    'plugged': int(r['plugged']),
                    'combo': int(r['combo']),
                    'gpon': int(r['gpon'])
                }
            print(f"Successfully loaded SFP data for {len(sfp_map)} OLTs from PonPort.")
        except Exception as e:
            print(f"Error loading PonPort SFP data: {e}")

    # 4. Load Board status file and build boards lists
    print("Loading Board status...")
    df_board = pd.read_excel(board_file)
    boards_map = {}
    for _, row in df_board.iterrows():
        dev_name = str(row.get('Device Name', '')).strip().upper()
        if not dev_name:
            continue
        slot = str(row.get('BoardName', '')).strip()
        b_type = str(row.get('BoardType', '')).strip()
        role = str(row.get('BoardRole', '')).strip()
        status = str(row.get('OperStatus', '')).strip()
        
        if pd.isna(row.get('BoardRole')) or role.lower() == 'nan':
            role = ""
        if pd.isna(row.get('BoardType')) or b_type.lower() == 'nan':
            b_type = ""
        if pd.isna(row.get('OperStatus')) or status.lower() == 'nan':
            status = ""
            
        # Filter: ONLY include PON Service Cards (exclude Ctrl, Power, Fan, Uplink, etc.)
        if not is_pon_service_card(b_type, role):
            continue
            
        card = {
            "slot": slot,
            "type": b_type,
            "role": role,
            "status": status
        }
        if dev_name not in boards_map:
            boards_map[dev_name] = []
        boards_map[dev_name].append(card)

    # Sort boards by slot position
    def slot_key(x):
        s = x['slot']
        parts = []
        for p in s.replace('/', '-').split('-'):
            try:
                parts.append(int(p))
            except ValueError:
                parts.append(p)
        return parts

    for dev in boards_map:
        boards_map[dev] = sorted(boards_map[dev], key=slot_key)

    # 5. Process OLT Location Excel
    print("Processing OLT Location...")
    df_loc = pd.read_excel(loc_file)
    new_data = {}

    for _, row in df_loc.iterrows():
        clli = str(row.get('CLLI_OLT', '')).strip().upper()
        if not clli or clli.lower() == 'nan':
            continue
            
        # Determine exact display name:
        if clli in device_master_names:
            display_name = clli
        else:
            alt_name = None
            if 'GO' in clli:
                cand = re.sub(r'GO([0-9])', r'G0\1', clli)
                if cand in device_master_names:
                    alt_name = cand
            elif 'G0' in clli:
                cand = re.sub(r'G0([0-9])', r'GO\1', clli)
                if cand in device_master_names:
                    alt_name = cand
            display_name = alt_name if alt_name else clli

        # Resolve Vendor from Device Master (polled from live NMS) first, fallback to df_loc
        vendor = get_prop(display_name, device_vendors)
        if not vendor:
            vendor = str(row.get('VENDOR_NAME', '')).strip().upper()
        
        model = get_prop(display_name, device_models, "")
            
        olt_type = str(row.get('OLT_TYPE', '')).strip().lower()
        if 'indoor' in olt_type:
            olt_type = 'indoor'
        elif 'outdoor' in olt_type:
            olt_type = 'outdoor'
        else:
            olt_type = get_prop(display_name, device_types, olt_type)
            
        # Hardcode / Enforce outdoor for OLTs verified in TPP Consolidate outdoor cabinets from photo audit
        if display_name in tpp_outdoor_olts or clli in tpp_outdoor_olts:
            olt_type = 'outdoor'
            
        ports = int(row.get('PORT_PON_TOTAL', 0)) if not pd.isna(row.get('PORT_PON_TOTAL')) else 0
        ports_use = int(row.get('PORT_PON_USE', 0)) if not pd.isna(row.get('PORT_PON_USE')) else 0
        
        # Override ports_use and ports total from OLT-Uplink-MaxSpeed if available
        p_total_override = get_prop(display_name, ports_total_map)
        p_use_override = get_prop(display_name, ports_use_map)
        
        if p_total_override is not None and p_total_override > 0:
            ports = p_total_override
            if p_use_override is not None:
                ports_use = p_use_override
        elif p_use_override is not None:
            ports_use = p_use_override
            
        ports_avail = max(0, ports - ports_use)
        
        # Count actual PON Service Cards installed from boards_map
        actual_boards = boards_map.get(display_name, boards_map.get(clli, []))
        if not actual_boards:
            if 'GO' in display_name:
                alt = re.sub(r'GO([0-9])', r'G0\1', display_name)
                if alt not in device_master_names or display_name not in device_master_names:
                    actual_boards = boards_map.get(alt, [])
            elif 'G0' in display_name:
                alt = re.sub(r'G0([0-9])', r'GO\1', display_name)
                if alt not in device_master_names or display_name not in device_master_names:
                    actual_boards = boards_map.get(alt, [])

        full_chassis_boards = build_chassis_service_slots(model, actual_boards)
        installed_pon_cards = len(actual_boards)
        cards_total = installed_pon_cards if installed_pon_cards > 0 else get_prop(display_name, slots_map, 0)
        
        edfa_cards = int(row.get('CARD_EDFA_TOTAL', 0)) if not pd.isna(row.get('CARD_EDFA_TOTAL')) else 0
        edfa_cnt_override = get_prop(display_name, edfa_count_map)
        if edfa_cards == 0 and edfa_cnt_override:
            edfa_cards = edfa_cnt_override
        
        lat = float(row.get('LATITUDE', 0.0)) if not pd.isna(row.get('LATITUDE')) else 0.0
        long_val = float(row.get('LONGITUDE', 0.0)) if not pd.isna(row.get('LONGITUDE')) else 0.0
        
        reg = str(row.get('REGION', '')).strip()
        prov = str(row.get('PROVINCE', '')).strip()
        dist = str(row.get('DISTRICT', '')).strip()
        subdist = str(row.get('SUBDISTRICT', '')).strip()
        
        hist = get_prop(display_name, historical_edfa, {'edfa_model': '', 'edfa_vendor': '', 'has_fwdm': False, 'fwdm_clli': ''})
        # Resolve EDFA directly from Latest EDFA Master Audit Report (Strict Ground Truth)
        if display_name in edfa_audited_olts or clli in edfa_audited_olts:
            edfa_model_val = get_prop(display_name, edfa_model_map, "")
            edfa_vendor_val = get_prop(display_name, edfa_vendor_map, "")
        else:
            edfa_model_val = hist.get('edfa_model', '')
            edfa_vendor_val = hist.get('edfa_vendor', '')
        
        has_fwdm = get_prop(display_name, fwdm_map, hist.get('has_fwdm', False))
        fwdm_clli_val = get_prop(display_name, fwdm_clli_map, hist.get('fwdm_clli', ''))
        
        # Resolve Uplink Site (Normalized Sitename e.g. PN1-PCT1291, DN1-CBR6043, CPE-ANC1602)
        raw_uplink_sites = get_prop(display_name, speed_uplink_sites, [])
        if not raw_uplink_sites:
            m_site = get_prop(display_name, mane_uplink_sites, "")
            if m_site:
                raw_uplink_sites = [m_site]
                
        cleaned_sites = []
        for r_s in raw_uplink_sites:
            norm_s = normalize_uplink_sitename(r_s)
            if norm_s and norm_s not in cleaned_sites:
                cleaned_sites.append(norm_s)
                
        # If both a prefixed site (e.g. PN1-PCT1291) and a plain site (e.g. PCT1291) exist, keep only the specific prefixed ones
        has_prefixed = any('-' in s for s in cleaned_sites)
        if has_prefixed:
            cleaned_sites = [s for s in cleaned_sites if '-' in s or not any(s in other for other in cleaned_sites if '-' in other)]
                
        uplink_site_val = ", ".join(cleaned_sites) if cleaned_sites else ""

        sfp_info = get_prop(display_name, sfp_map, {})
        
        new_data[display_name] = {
            "reg": reg,
            "prov": prov,
            "dist": dist,
            "subdist": subdist,
            "lat": lat,
            "long": long_val,
            "vendor": vendor,
            "model": model,
            "type": olt_type,
            "ip": get_prop(display_name, device_ips, get_prop(display_name, speed_ip_map, "")),
            "uplink_site": uplink_site_val,
            "cards": cards_total,
            "edfa_cards": edfa_cards,
            "ports": ports,
            "ports_use": ports_use,
            "ports_avail": ports_avail,
            "sfp_plugged": sfp_info.get('plugged', None),
            "sfp_combo": sfp_info.get('combo', 0),
            "sfp_gpon": sfp_info.get('gpon', 0),
            "boards": full_chassis_boards,
            "edfa_model": edfa_model_val,
            "edfa_vendor": edfa_vendor_val,
            "has_fwdm": has_fwdm,
            "fwdm_clli": fwdm_clli_val,
            "uplink_duplex": get_prop(display_name, speed_map, "-"),
            "onu_count": get_prop(display_name, onu_count_map, 0),
            "corp_links": get_prop(display_name, corp_map, 0),
            "put_inservice": get_prop(display_name, inservice_map, "")
        }

    # 6. Save new JSON & Embed into HTML across all target directories
    # Determine as_of date from latest date among all input files
    all_input_files = [edfa_file, loc_file, speed_file, device_file, board_file, corp_file, mane_file, pon_file]
    dates = [extract_date_from_filename(f) for f in all_input_files if f]
    dates = [d for d in dates if d]
    latest_dt = max(dates) if dates else datetime.now()
    as_of_str = latest_dt.strftime('%d %b %Y')
    as_of_badge = latest_dt.strftime('%d/%m/%y')

    json_str = json.dumps(new_data, ensure_ascii=False)

    for tdir in TARGET_DIRS:
        if not os.path.exists(tdir):
            continue
        
        # 6.1 Save JSON
        jpath = os.path.join(tdir, 'olt_location_data.json')
        print(f"Saving updated data to {jpath}...")
        try:
            with open(jpath, 'w', encoding='utf-8') as f:
                f.write(json_str)
        except Exception as e:
            print(f"Error saving to {jpath}: {e}")

        # 6.2 Embed in HTML
        hpath = os.path.join(tdir, 'index.html')
        if os.path.exists(hpath):
            print(f"Embedding in {hpath}...")
            try:
                new_lines = []
                replaced_data = False
                with open(hpath, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith('const locationData ='):
                            new_lines.append(f'const locationData = {json_str};\n')
                            replaced_data = True
                        else:
                            if 'id="asOfDate"' in line or 'class="badge-date"' in line:
                                line = re.sub(r'(<span[^>]*id="asOfDate"[^>]*>)[^<]*(</span>)',
                                              f'\\g<1>{as_of_str}\\g<2>', line)
                                line = re.sub(r'(<span[^>]*class="badge-date"[^>]*>)[^<]*(</span>)',
                                              f'\\g<1>{as_of_str}\\g<2>', line)
                            if '📅 AS OF' in line:
                                line = re.sub(r'📅 AS OF \d{2}/\d{2}/\d{2,4}', f'📅 AS OF {as_of_badge}', line)
                            if 'class="panel-asof"' in line:
                                line = re.sub(r'\(as of [^\)]*\)', f'(as of {as_of_str})', line)
                            new_lines.append(line)

                if replaced_data:
                    with open(hpath, 'w', encoding='utf-8') as f:
                        f.writelines(new_lines)
                    print(f"Successfully updated embedded locationData and AS OF date ({as_of_badge} / {as_of_str}) in {hpath}")
                else:
                    print(f"Warning: Could not find 'const locationData =' line in {hpath}")
            except Exception as e:
                print(f"Error updating {hpath}: {e}")

    print("Update Completed Successfully!")

if __name__ == '__main__':
    main()
