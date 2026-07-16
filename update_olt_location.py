import os
import glob
import pandas as pd
import json

# Define paths
BASE_DIR = '/Users/bbae/GPTCodex'
LOCATION_DIR = '/Users/bbae/OLT-Location'
JSON_PATH = os.path.join(LOCATION_DIR, 'olt_location_data.json')
HTML_PATH = os.path.join(LOCATION_DIR, 'index.html')

def main():
    # 1. Load historical EDFA models from existing JSON
    historical_edfa = {}
    if os.path.exists(JSON_PATH):
        print("Loading historical EDFA mappings...")
        try:
            with open(JSON_PATH, 'r', encoding='utf-8') as f:
                old_data = json.load(f)
                for k, v in old_data.items():
                    if v.get('edfa_model') or v.get('edfa_vendor'):
                        historical_edfa[k] = {
                            'edfa_model': v.get('edfa_model', ''),
                            'edfa_vendor': v.get('edfa_vendor', '')
                        }
            print(f"Loaded {len(historical_edfa)} historical EDFA records.")
        except Exception as e:
            print(f"Error loading historical JSON: {e}")

    # 2. Identify latest files
    loc_file = os.path.join(BASE_DIR, 'reports', 'inventory_port_olt_20260713.xlsx')
    if not os.path.exists(loc_file):
        # Fallback to the latest in reports/
        files = sorted(glob.glob(os.path.join(BASE_DIR, 'reports', 'inventory_port_olt_*.xlsx')), key=os.path.getmtime, reverse=True)
        if files:
            loc_file = files[0]
        else:
            # Fallback to data/misc
            files = sorted(glob.glob(os.path.join(BASE_DIR, 'data', 'misc', 'OLT_Location*.xlsx')), key=os.path.getmtime, reverse=True)
            loc_file = files[0] if files else None

    # Find device and board files
    device_files = sorted(glob.glob(os.path.join(BASE_DIR, 'data', 'device', 'Device-*.xlsx')), key=os.path.getmtime, reverse=True)
    device_file = device_files[0] if device_files else None

    board_files = sorted(glob.glob(os.path.join(BASE_DIR, 'cardolt', 'Board-*.xlsx')), key=os.path.getmtime, reverse=True)
    board_file = board_files[0] if board_files else None

    if not loc_file or not device_file or not board_file:
        print("Error: Missing required files.")
        return

    print(f"Using location file: {loc_file}")
    print(f"Using device file:   {device_file}")
    print(f"Using board file:    {board_file}")

    # 3. Load Device Master to map OLT name -> Model & OLT Type
    print("Loading Device master...")
    df_dev = pd.read_excel(device_file)
    device_models = {}
    device_types = {}
    for _, row in df_dev.iterrows():
        name = str(row.get('Device Name', '')).strip().upper()
        model = str(row.get('Model', '')).strip()
        typ = str(row.get('OLT Type', '')).strip().lower()
        if name:
            device_models[name] = model
            device_types[name] = typ

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
            
        clli_std = clli.replace('GO0', 'G00')
        
        vendor = str(row.get('VENDOR_NAME', '')).strip().upper()
        
        model = device_models.get(clli_std, "")
        if not model:
            model = ""
            
        olt_type = str(row.get('OLT_TYPE', '')).strip().lower()
        if 'indoor' in olt_type:
            olt_type = 'indoor'
        elif 'outdoor' in olt_type:
            olt_type = 'outdoor'
        else:
            olt_type = device_types.get(clli_std, olt_type)
            
        ports = int(row.get('PORT_PON_TOTAL', 0)) if not pd.isna(row.get('PORT_PON_TOTAL')) else 0
        ports_use = int(row.get('PORT_PON_USE', 0)) if not pd.isna(row.get('PORT_PON_USE')) else 0
        ports_avail = int(row.get('PORT_PON_AVAILABLE', 0)) if not pd.isna(row.get('PORT_PON_AVAILABLE')) else 0
        
        cards_total = int(row.get('CARD_OLT_TOTAL', 0)) if not pd.isna(row.get('CARD_OLT_TOTAL')) else 0
        edfa_cards = int(row.get('CARD_EDFA_TOTAL', 0)) if not pd.isna(row.get('CARD_EDFA_TOTAL')) else 0
        
        lat = float(row.get('LATITUDE', 0.0)) if not pd.isna(row.get('LATITUDE')) else 0.0
        long_val = float(row.get('LONGITUDE', 0.0)) if not pd.isna(row.get('LONGITUDE')) else 0.0
        
        reg = str(row.get('REGION', '')).strip()
        prov = str(row.get('PROVINCE', '')).strip()
        dist = str(row.get('DISTRICT', '')).strip()
        subdist = str(row.get('SUBDISTRICT', '')).strip()
        
        hist = historical_edfa.get(clli_std, historical_edfa.get(clli, {'edfa_model': '', 'edfa_vendor': ''}))
        
        new_data[clli_std] = {
            "reg": reg,
            "prov": prov,
            "dist": dist,
            "subdist": subdist,
            "lat": lat,
            "long": long_val,
            "vendor": vendor,
            "model": model,
            "type": olt_type,
            "cards": cards_total,
            "edfa_cards": edfa_cards,
            "ports": ports,
            "ports_use": ports_use,
            "ports_avail": ports_avail,
            "boards": boards_map.get(clli_std, boards_map.get(clli, [])),
            "edfa_model": hist['edfa_model'],
            "edfa_vendor": hist['edfa_vendor']
        }

    # 6. Save new JSON
    print(f"Saving updated data to {JSON_PATH}...")
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(new_data, f, ensure_ascii=False)

    # 7. Embed in HTML
    print(f"Embedding in {HTML_PATH}...")
    if os.path.exists(HTML_PATH):
        json_str = json.dumps(new_data, ensure_ascii=False)
        
        new_lines = []
        replaced = False
        with open(HTML_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                if line.startswith('const locationData ='):
                    new_lines.append(f'const locationData = {json_str};\n')
                    replaced = True
                else:
                    new_lines.append(line)
                    
        if replaced:
            with open(HTML_PATH, 'w', encoding='utf-8') as f:
                f.writelines(new_lines)
            print("Successfully updated embedded locationData in HTML.")
        else:
            print("Error: Could not find 'const locationData =' line in HTML.")
                
    print("Update Completed Successfully!")

if __name__ == '__main__':
    main()
