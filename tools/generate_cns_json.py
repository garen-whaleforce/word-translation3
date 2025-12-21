# tools/generate_cns_json.py
"""
從 CB PDF 抽取的原始資料生成 CNS 15598-1 JSON
使用 cb_overview_raw.json 和 cb_clauses_raw.json
"""
import json
import re
import argparse
from pathlib import Path

def load_json(path: Path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def extract_meta_from_chunks(chunks: list, pdf_name: str) -> dict:
    """從文字 chunks 中抽取 meta 資訊"""
    meta = {
        "source_pdf_name": pdf_name,
        "standard": "IEC 62368-1:2018",
        "target_report": "CNS 15598-1 (109年版)",
        "cb_report_no": "",
        "report_date": "",
        "applicant": "",
        "manufacturer": "",
        "factory_locations": [],
        "trade_mark": "",
        "model_type_references": [],
        "ratings_input": "",
        "ratings_output": "",
        "mass_of_equipment": "",
        "notes": [],
        # Test item particulars 欄位
        "test_item_particulars": {
            "product_group": "",  # end product / built-in component
            "classification_of_use": [],  # Ordinary person, Instructed person, Skilled person, Children likely present
            "supply_connection": "",  # AC mains / DC mains / not mains connected
            "supply_tolerance": "",  # +10%/-10%, +20%/-15%, etc.
            "supply_connection_type": "",  # pluggable equipment type A/B, direct plug-in, etc.
            "protective_device_rating": "",  # 16A, 20A, etc.
            "equipment_mobility": [],  # movable, hand-held, transportable, direct plug-in, stationary, etc.
            "overvoltage_category": "",  # OVC I, OVC II, OVC III, OVC IV
            "equipment_class": "",  # Class I, Class II, Class III
            "special_installation_location": "",  # N/A, restricted access area, outdoor location
            "pollution_degree": "",  # PD 1, PD 2, PD 3
            "manufacturer_tma": "",  # 45 °C, etc.
            "ip_protection_class": "",  # IPX0, IP__
            "power_systems": "",  # TN, TT, IT
            "altitude_operation": "",  # 2000 m or less, 5000 m
            "altitude_test_lab": "",  # 2000 m or less
        }
    }

    # 合併前幾頁文字
    first_pages_text = '\n'.join([c['text'] for c in chunks[:15]])

    # Report No - 多種格式
    m = re.search(r'Report\s*Number[.\s]*:\s*([A-Z0-9]+\s*\d+)', first_pages_text, re.IGNORECASE)
    if m:
        meta['cb_report_no'] = m.group(1).strip()
    else:
        m = re.search(r'Report\s*No\.?\s*[:\s]*([A-Z0-9]{2,}\s*\d+)', first_pages_text, re.IGNORECASE)
        if m:
            meta['cb_report_no'] = m.group(1).strip()

    # Date
    m = re.search(r'Date\s*of\s*issue\s*[.\s]*:\s*(\d{4}[-/]\d{2}[-/]\d{2})', first_pages_text, re.IGNORECASE)
    if m:
        meta['report_date'] = m.group(1)

    # Applicant - 多種格式
    m = re.search(r"Applicant.*?:\s*([A-Z][^\n]+)", first_pages_text, re.IGNORECASE)
    if m:
        meta['applicant'] = m.group(1).strip()

    # Manufacturer - 多種格式
    m = re.search(r'Manufacturer\s*[.\s]*:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if m:
        mfr = m.group(1).strip()
        if 'same as' in mfr.lower() or 'see above' in mfr.lower():
            meta['manufacturer'] = 'Same as applicant'
        else:
            meta['manufacturer'] = mfr

    # Model - 多種格式
    m = re.search(r'Model/Type\s*reference\s*[.\s]*:\s*([A-Z0-9][\w\-]+(?:\s*,\s*[A-Z0-9][\w\-]+)?)', first_pages_text, re.IGNORECASE)
    if m:
        models = m.group(1).strip()
        # 分割多個型號（只用逗號分隔）
        model_list = re.split(r'\s*,\s*', models)
        meta['model_type_references'] = [m.strip() for m in model_list if m.strip()]
    else:
        m = re.search(r'Model[/\s]*Type\s*Ref[:\s]*([^\n]+)', first_pages_text, re.IGNORECASE)
        if m:
            models = m.group(1).strip()
            model_list = re.split(r'\s*,\s*', models)
            meta['model_type_references'] = [m.strip() for m in model_list if m.strip()]

    # Ratings - 多種格式
    m = re.search(r'Ratings\s*[.\s]*:\s*Input:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if m:
        meta['ratings_input'] = m.group(1).strip()
    else:
        m = re.search(r'Rated\s*input[:\s]*([^\n]+)', first_pages_text, re.IGNORECASE)
        if m:
            meta['ratings_input'] = m.group(1).strip()
        else:
            m = re.search(r'Input:\s*([0-9\-]+V[^\n]+)', first_pages_text, re.IGNORECASE)
            if m:
                meta['ratings_input'] = m.group(1).strip()

    # Ratings Output - 抓取多行直到遇到下一個欄位標題或空行
    m = re.search(r'Output[:\s]*(.*?)(?=\n[A-Z][a-z]+[:\s]|\n\n|\nTest|\nNotes)', first_pages_text, re.IGNORECASE | re.DOTALL)
    if m:
        # 將多行合併為單行，去除多餘空白
        output_text = m.group(1).strip()
        output_text = re.sub(r'\s*\n\s*', ' ', output_text)  # 換行轉空格
        output_text = re.sub(r'\s+', ' ', output_text)  # 多空格合併
        meta['ratings_output'] = output_text
    else:
        m = re.search(r'Rated\s*output[:\s]*(.*?)(?=\n[A-Z][a-z]+[:\s]|\n\n)', first_pages_text, re.IGNORECASE | re.DOTALL)
        if m:
            output_text = m.group(1).strip()
            output_text = re.sub(r'\s*\n\s*', ' ', output_text)
            output_text = re.sub(r'\s+', ' ', output_text)
            meta['ratings_output'] = output_text

    # Mass of equipment - 設備質量
    # 格式可能是：
    # 1. "Mass of equipment (kg) ...: Approx. 0.072kg."
    # 2. "Mass of equipment (kg) ...: For direct plug-in models approx. 0.134kg;\nFor desktop models approx. 0.135Kg."
    m = re.search(r'Mass\s+of\s+equipment\s*\(kg\)\s*[.\s]*:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if m:
        mass_text = m.group(1).strip()
        # 如果結尾是分號，可能有多行
        if mass_text.endswith(';'):
            # 嘗試抓取下一行
            m2 = re.search(r'Mass\s+of\s+equipment\s*\(kg\)\s*[.\s]*:\s*([^\n]+\n[^\n]+)', first_pages_text, re.IGNORECASE)
            if m2:
                mass_text = m2.group(1).strip()
                mass_text = re.sub(r'\s*\n\s*', ' ', mass_text)
        # 移除可能誤抓的 TRF No. 等文字
        mass_text = re.sub(r'\s*TRF\s+No\..*$', '', mass_text, flags=re.IGNORECASE)
        mass_text = re.sub(r'\s+', ' ', mass_text).strip()
        meta['mass_of_equipment'] = mass_text

    # ===== Test item particulars 欄位提取 =====
    tip = meta['test_item_particulars']

    # Product group - 產品群組
    m = re.search(r'Product\s+group\s*[.\s]*:\s*(end\s+product|built-in\s+component)', first_pages_text, re.IGNORECASE)
    if m:
        tip['product_group'] = m.group(1).strip().lower()

    # Classification of use - 使用分類
    # 檢查哪些選項被選中（通常 PDF 中選中的項目會出現在特定位置）
    classification_section = re.search(r'Classification\s+of\s+use\s+by\s*[.\s]*:\s*([^\n]+(?:\n[^\n]+){0,3})', first_pages_text, re.IGNORECASE)
    if classification_section:
        section_text = classification_section.group(1)
        classifications = []
        if 'Ordinary' in section_text:
            classifications.append('ordinary')
        if 'Instructed' in section_text:
            classifications.append('instructed')
        if 'Skilled' in section_text:
            classifications.append('skilled')
        if 'Children' in section_text:
            classifications.append('children')
        tip['classification_of_use'] = classifications

    # Supply connection - 電源連接
    m = re.search(r'Supply\s+connection\s*[.\s]*:\s*(AC\s+mains|DC\s+mains|not\s+mains\s+connected)', first_pages_text, re.IGNORECASE)
    if m:
        tip['supply_connection'] = m.group(1).strip().lower()

    # Supply tolerance - 電源許可差
    m = re.search(r'Supply\s+tolerance\s*[.\s]*:\s*(\+\d+%/-\d+%)', first_pages_text, re.IGNORECASE)
    if m:
        tip['supply_tolerance'] = m.group(1).strip()

    # Equipment mobility - 設備移動性 (重要！需要精確識別)
    # PDF 格式: "Equipment mobility ..........: movable hand-held transportable\ndirect plug-in stationary for building-in"
    mobility_section = re.search(r'Equipment\s+mobility\s*[.\s]*:\s*([^\n]+(?:\n[^\n]+)?)', first_pages_text, re.IGNORECASE)
    if mobility_section:
        section_text = mobility_section.group(1).lower()
        mobility_options = []
        # 根據 PDF 文字識別選中的選項
        # 注意：PDF 中通常只有被選中的項目會顯示，或者位置決定
        if 'movable' in section_text and 'hand-held' not in section_text.split('movable')[0][-20:]:
            mobility_options.append('movable')
        if 'hand-held' in section_text:
            mobility_options.append('hand-held')
        if 'transportable' in section_text:
            mobility_options.append('transportable')
        if 'direct plug-in' in section_text or 'direct plug' in section_text:
            mobility_options.append('direct plug-in')
        if 'stationary' in section_text:
            mobility_options.append('stationary')
        if 'for building-in' in section_text or 'building-in' in section_text:
            mobility_options.append('for building-in')
        if 'wall' in section_text or 'ceiling' in section_text:
            mobility_options.append('wall/ceiling-mounted')
        if 'srme' in section_text or 'rack-mounted' in section_text:
            mobility_options.append('SRME/rack-mounted')
        tip['equipment_mobility'] = mobility_options

    # Overvoltage category - 過壓類別
    m = re.search(r'Overvoltage\s+category\s*\(OVC\)\s*[.\s]*:\s*(OVC\s*[IV]+)', first_pages_text, re.IGNORECASE)
    if m:
        tip['overvoltage_category'] = m.group(1).strip().upper().replace(' ', '')

    # Class of equipment - 防電擊保護
    m = re.search(r'Class\s+of\s+equipment\s*[.\s]*:\s*(Class\s*[I]+|Not\s+classified)', first_pages_text, re.IGNORECASE)
    if m:
        tip['equipment_class'] = m.group(1).strip()

    # Special installation location - 特殊安裝位置
    m = re.search(r'Special\s+installation\s+location\s*[.\s]*:\s*(N/A|restricted\s+access\s+area|outdoor\s+location)', first_pages_text, re.IGNORECASE)
    if m:
        tip['special_installation_location'] = m.group(1).strip()

    # Pollution degree - 污染等級
    m = re.search(r'Pollution\s+degree\s*\(PD\)\s*[.\s]*:\s*(PD\s*[123])', first_pages_text, re.IGNORECASE)
    if m:
        tip['pollution_degree'] = m.group(1).strip().upper().replace(' ', '')

    # Manufacturer's specified Tma - 製造商宣告 Tma
    m = re.search(r"Manufacturer.*?T\s*[.\s]*:\s*(\d+)\s*°?C", first_pages_text, re.IGNORECASE)
    if m:
        tip['manufacturer_tma'] = m.group(1) + " °C"

    # IP protection class - IP 等級
    m = re.search(r'IP\s+protection\s+class\s*[.\s]*:\s*(IP[X0-9]+)', first_pages_text, re.IGNORECASE)
    if m:
        tip['ip_protection_class'] = m.group(1).strip().upper()

    # Power systems - 電力系統
    power_section = re.search(r'Power\s+systems\s*[.\s]*:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if power_section:
        section_text = power_section.group(1).upper()
        systems = []
        if 'TN' in section_text:
            systems.append('TN')
        if 'TT' in section_text:
            systems.append('TT')
        if 'IT' in section_text:
            systems.append('IT')
        if 'NOT AC MAINS' in section_text:
            systems.append('not AC mains')
        tip['power_systems'] = ', '.join(systems) if systems else section_text.strip()

    # Altitude during operation - 設備適用的海拔高度
    m = re.search(r'Altitude\s+during\s+operation\s*\(m\)\s*[.\s]*:\s*(2000\s*m\s+or\s+less|5000\s*m)', first_pages_text, re.IGNORECASE)
    if m:
        tip['altitude_operation'] = m.group(1).strip()

    # Altitude of test laboratory - 測試實驗室海拔高度
    m = re.search(r'Altitude\s+of\s+test\s+laboratory\s*\(m\)\s*[.\s]*:\s*(2000\s*m\s+or\s+less|\d+\s*m)', first_pages_text, re.IGNORECASE)
    if m:
        tip['altitude_test_lab'] = m.group(1).strip()

    # Protective device rating - 保護裝置額定電流
    m = re.search(r'Considered\s+current\s+rating.*?:\s*(\d+)\s*A', first_pages_text, re.IGNORECASE | re.DOTALL)
    if m:
        tip['protective_device_rating'] = m.group(1) + " A"

    # ===== 備註區塊提取 (General product information and other remarks) =====
    # 這些備註會填入 Word 模板的 T4R19

    # General product information and other remarks
    m = re.search(r'General\s+product\s+information\s+and\s+other\s+remarks:\s*(.*?)(?=Model\s+Differences|Additional\s+application|TRF\s+No)', first_pages_text, re.IGNORECASE | re.DOTALL)
    if m:
        remarks_text = m.group(1).strip()
        # 清理頁碼和報告號碼
        remarks_text = re.sub(r'TRF\s+No\.\s+IEC62368_1E', '', remarks_text)
        remarks_text = re.sub(r'Page\s+\d+\s+of\s+\d+\s+Report\s+No\.\s+[\w-]+', '', remarks_text)
        remarks_text = re.sub(r'\s+', ' ', remarks_text).strip()
        meta['general_product_remarks'] = remarks_text

    # Model Differences
    m = re.search(r'Model\s+Differences:\s*(.*?)(?=Additional\s+application|TRF\s+No|$)', first_pages_text, re.IGNORECASE | re.DOTALL)
    if m:
        model_diff = m.group(1).strip()
        model_diff = re.sub(r'TRF\s+No\.\s+IEC62368_1E', '', model_diff)
        model_diff = re.sub(r'\s+', ' ', model_diff).strip()
        meta['model_differences'] = model_diff

    # Name and address of factory - 工廠資訊
    m = re.search(r'Name\s+and\s+address\s+of\s+factory.*?:\s*(.*?)(?=General\s+product\s+information)', first_pages_text, re.IGNORECASE | re.DOTALL)
    if m:
        factory_text = m.group(1).strip()
        # 提取工廠列表
        factories = re.findall(r'\d+\)\s*([^\d]+?)(?=\d+\)|$)', factory_text, re.DOTALL)
        if factories:
            meta['factory_locations'] = [f.strip() for f in factories if f.strip()]

    return meta

def convert_overview_to_cns(overview_raw: list) -> list:
    """將原始 overview 資料轉換為 CNS 格式"""
    result = []

    clause_map = {
        '5': 'Clause 5 Electrically-caused injury',
        '6': 'Clause 6 Electrically-caused fire',
        '7': 'Clause 7 Injury caused by hazardous substances',
        '8': 'Clause 8 Mechanically-caused injury',
        '9': 'Clause 9 Thermal burn',
        '10': 'Clause 10 Radiation'
    }

    for item in overview_raw:
        clause = item.get('clause', '')
        row = item.get('row', [])

        if len(row) < 5:
            continue

        energy_source = row[0].replace('\n', ' ').strip()
        parts_involved = row[1].replace('\n', ' ').strip()

        # 組合 safeguards (B, S, R 或 B, 1st S, 2nd S)
        safeguards_parts = []
        for i, label in enumerate(['B', 'S', 'R']):
            if i + 2 < len(row) and row[i + 2] and row[i + 2] != 'N/A':
                val = row[i + 2].replace('\n', ' ').strip()
                if val and val != 'N/A':
                    safeguards_parts.append(f"{label}: {val}")

        safeguards = ', '.join(safeguards_parts) if safeguards_parts else 'N/A'

        # 抽取 energy source class (ES3, PS2, MS1 等)
        energy_class = ''
        m = re.match(r'^(ES[123]|PS[123]|MS[123]|TS[123]|RS[123]|N/A)', energy_source)
        if m:
            energy_class = m.group(1)

        result.append({
            'energy_source_class': energy_class,
            'parts_involved': energy_source,
            'safeguards': safeguards,
            'remarks_or_clause_ref': clause_map.get(clause, f'Clause {clause}'),
            'evidence_quote': ' '.join([c.replace('\n', ' ') for c in row])
        })

    return result

def dedupe_clauses(clauses_raw: list) -> list:
    """去重 clause（保留第一個出現的）"""
    seen = set()
    result = []
    for c in clauses_raw:
        cid = c.get('clause_id', '')
        if cid and cid not in seen:
            seen.add(cid)
            result.append(c)
    return result

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input_dir", required=True, help="包含 cb_*.json 的目錄")
    ap.add_argument("--pdf_name", required=True, help="PDF 檔名")
    ap.add_argument("--out", required=True, help="輸出 JSON 路徑")
    ap.add_argument("--special_tables", default=None, help="特殊表格 JSON 路徑 (cb_special_tables.json)")
    args = ap.parse_args()

    input_dir = Path(args.input_dir)

    # 讀取原始資料
    chunks = load_json(input_dir / "cb_text_chunks.json")
    overview_raw = load_json(input_dir / "cb_overview_raw.json")
    clauses_raw = load_json(input_dir / "cb_clauses_raw.json")

    # 讀取特殊表格（如果有）
    special_tables = {}
    special_tables_path = Path(args.special_tables) if args.special_tables else input_dir / "cb_special_tables.json"
    if special_tables_path.exists():
        special_tables = load_json(special_tables_path)

    # 生成各區塊
    meta = extract_meta_from_chunks(chunks, args.pdf_name)
    overview = convert_overview_to_cns(overview_raw)
    clauses = dedupe_clauses(clauses_raw)

    # 從特殊表格抽取 overview_cb_p12_rows
    overview_cb_p12_rows = []
    if 'overview' in special_tables and 'rows' in special_tables['overview']:
        overview_cb_p12_rows = special_tables['overview']['rows']

    # 組合最終 JSON
    result = {
        'meta': meta,
        'overview_energy_sources_and_safeguards': overview,
        'overview_cb_p12_rows': overview_cb_p12_rows,
        'clauses': clauses,
        'attachments_or_annex': []
    }

    # 輸出
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Generated: {out_path}")
    print(f"overview_cb_p12_rows: {len(overview_cb_p12_rows)} rows")
    print(f"clauses: {len(clauses)} rows")

if __name__ == "__main__":
    main()
