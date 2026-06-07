import streamlit as st
import fitz
import re
import io
import copy
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from datetime import date
import os

st.set_page_config(page_title="ไทวัสดุ — คำนวณกล่องสินค้า", page_icon="📦", layout="wide")

st.markdown("""
<style>
div[data-testid="stSidebar"] { background: #1a1a2e; }
div[data-testid="stSidebar"] * { color: #eee !important; }
.brand { background: linear-gradient(135deg,#c62828,#e53935); color:#fff; padding:1.2rem 1.5rem;
         border-radius:12px; margin-bottom:1.5rem; }
.brand h1 { margin:0; font-size:1.6rem; }
.brand p  { margin:0; opacity:.85; font-size:.9rem; }
.stat-box { background:#fff; border:1px solid #e0e0e0; border-radius:10px;
            padding:1rem; text-align:center; }
.stat-box .num { font-size:2rem; font-weight:700; color:#c62828; }
.stat-box .lbl { font-size:.8rem; color:#888; }
.branch-card { background:#fff; border:1px solid #e8e8e8; border-radius:10px;
               padding:1rem 1.2rem; margin-bottom:.6rem; }
.badge { display:inline-block; padding:2px 8px; border-radius:20px;
         font-size:.75rem; font-weight:600; margin:1px; }
.b-mixed  { background:#fff8e1; color:#f57f17; border:1px solid #ffe082; }
.b-foam   { background:#e8f5e9; color:#2e7d32; border:1px solid #a5d6a7; }
.b-canvas { background:#ffebee; color:#c62828; border:1px solid #ef9a9a; }
.b-num    { background:#e3f2fd; color:#1565c0; border:1px solid #90caf9; }
.ok   { color:#2e7d32; font-weight:600; }
.warn { color:#c62828; font-weight:600; }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "branches.xlsx")

@st.cache_data
def load_branch_master():
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["Store รหัสสาขา"]
    data = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row[2] and str(row[2]).strip().isdigit():
            upfront = int(row[2])
            data[upfront] = {
                "name_en": (row[3] or "").strip(),
                "name_th": (row[8] or "").strip(),
            }
    return data

BRANCH_MASTER = load_branch_master()

# ─── Box calculation ───────────────────────────────────────────────────────────
def calc_boxes(canvas_pairs: int, foam_dozen: float) -> list[dict]:
    boxes, cp, fd = [], canvas_pairs, foam_dozen
    mixed = min(int(fd), cp // 6)
    for _ in range(mixed):
        boxes.append({"type": "mixed", "label": "ฟองน้ำ 1 โหล + ผ้าใบ 6 คู่"})
        fd -= 1; cp -= 6
    while fd >= 2:
        boxes.append({"type": "foam", "label": "ฟองน้ำ 2 โหล"})
        fd -= 2
    if fd >= 1:
        if cp >= 6:
            boxes.append({"type": "mixed", "label": "ฟองน้ำ 1 โหล + ผ้าใบ 6 คู่"})
            cp -= 6
        else:
            boxes.append({"type": "foam", "label": "ฟองน้ำ 1 โหล"})
        fd = 0
    while cp >= 12:
        boxes.append({"type": "canvas", "label": "ผ้าใบ 12 คู่"})
        cp -= 12
    if cp > 0:
        boxes.append({"type": "canvas", "label": f"ผ้าใบ {cp} คู่"})
    return boxes

# ─── PDF Parser ────────────────────────────────────────────────────────────────
def parse_pdf(pdf_bytes: bytes, filename: str) -> dict:
    """
    อ่าน PDF ใบแบ่งสินค้าไทวัสดุ (text-based)
    คืนค่า: {
        'po_no': str,
        'po_items': {'canvas': int, 'foam200': int, 'foam212': int},   # จากใบสั่งซื้อ
        'branches': {code: {'name':str,'canvas':int,'foam200':int,'foam212':int}},
        'errors': [str]
    }
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    result = {"po_no": "", "po_items": {"canvas": 0, "foam200": 0, "foam212": 0},
              "branches": {}, "errors": [], "filename": filename}

    # หาเลข PO จากหน้าแรก
    first_text = doc[0].get_text()
    m = re.search(r'SO\d+-\d+', first_text)
    if m:
        result["po_no"] = m.group(0)

    # ─ ใบสั่งซื้อ: หน้าที่มี "ใบสสงซซอ" หรือ "ใบสั่งซื้อ" ─
    # นับ qty EACH จาก column จำนวน(หน่วยขาย)
    po_pages = []
    dist_pages = []

    for i in range(len(doc)):
        text = doc[i].get_text()
        if "ใบแบงสสนคคา" in text or "ใบแบบงสสนคคา" in text or "กระจายไปสาขา" in text:
            dist_pages.append(i)
        elif "ใบสสงซซอ" in text or "จสานวนสสง" in text:
            po_pages.append(i)

    # Parse ใบสั่งซื้อ
    for pi in po_pages:
        text = doc[pi].get_text()
        lines = text.split('\n')
        for j, line in enumerate(lines):
            # หาบรรทัดที่มีชื่อสินค้า
            if 'รองเทคาผคาใบ' in line or 'NANYANG 205' in line:
                # หา qty EACH ใน context
                ctx = '\n'.join(lines[j:j+15])
                for m2 in re.finditer(r'(\d+\.\d+)\s*\n?\s*EACH', ctx):
                    qty = float(m2.group(1))
                    if qty >= 6:  # qty หน่วยขาย (ไม่ใช่ pack)
                        result["po_items"]["canvas"] += qty
                    break

            elif 'รองเทคาแตะ' in line or 'หหนทบ' in line or 'สวม' in line:
                ctx = '\n'.join(lines[j:j+15])
                # หา รุ่น จาก context
                is_212 = '212' in '\n'.join(lines[max(0,j-2):j+5]) or 'สวม' in line
                is_200 = '200' in '\n'.join(lines[max(0,j-2):j+5]) or 'หหนทบ' in line
                for m2 in re.finditer(r'(\d+\.\d+)\s*\n?\s*EACH', ctx):
                    qty = float(m2.group(1))
                    if qty >= 12:
                        if is_212:
                            result["po_items"]["foam212"] += qty
                        elif is_200:
                            result["po_items"]["foam200"] += qty
                    break

    # Parse ใบแบ่งสินค้า
    current_branch = None
    for pi in dist_pages:
        text = doc[pi].get_text()
        lines = text.split('\n')

        for j, line in enumerate(lines):
            # หาสาขา
            bm = re.search(r'กระจายไปสาขา\s+(\d{5})\s+(.*)', line)
            if bm:
                code = int(bm.group(1))
                name_raw = bm.group(2).strip()
                current_branch = code
                if code not in result["branches"]:
                    # lookup master
                    master = BRANCH_MASTER.get(code, {})
                    result["branches"][code] = {
                        "name": master.get("name_en") or name_raw,
                        "name_th": master.get("name_th") or "",
                        "canvas": 0, "foam200": 0, "foam212": 0,
                    }

            # หาจำนวนสินค้า: บรรทัดตัวเลขลอยๆ เช่น " 12.00"
            if current_branch and re.match(r'^\s*\d+\.\d+\s*$', line):
                qty = float(line.strip())
                if qty < 3:
                    continue  # น้อยเกิน ข้ามไป

                # ดู context ก่อนหน้า
                ctx = ' '.join(lines[max(0, j-8):j])
                if '212' in ctx or ('สวม' in ctx and '200' not in ctx):
                    result["branches"][current_branch]["foam212"] += qty
                elif '200' in ctx or 'หหนทบ' in ctx or 'แตะ' in ctx:
                    result["branches"][current_branch]["foam200"] += qty
                elif 'ผคาใบ' in ctx or '205' in ctx:
                    result["branches"][current_branch]["canvas"] += qty

    doc.close()
    return result

def verify_po_vs_dist(po_items: dict, branches: dict) -> dict:
    """เปรียบเทียบยอดใบสั่งซื้อ vs ใบแบ่งสินค้า"""
    dist = {"canvas": 0, "foam200": 0, "foam212": 0}
    for b in branches.values():
        dist["canvas"] += b["canvas"]
        dist["foam200"] += b["foam200"]
        dist["foam212"] += b["foam212"]

    result = {}
    for key in ["canvas", "foam200", "foam212"]:
        diff = po_items[key] - dist[key]
        result[key] = {"po": po_items[key], "dist": dist[key], "diff": diff, "ok": diff == 0}
    return result

# ─── Excel generation ──────────────────────────────────────────────────────────
def generate_excel(all_branches: list[dict], company: str, invoice_no: str,
                   invoice_date, total_boxes: int) -> bytes:
    wb_tpl = load_workbook(TEMPLATE_PATH)
    tpl_ws = wb_tpl["ใบปะหน้า"]
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    # Sheet สรุป
    ws_sum = wb_out.create_sheet("สรุปทุกสาขา")
    headers = ["ลำดับ", "สาขา (EN)", "สาขา (TH)", "รหัส",
               "ผ้าใบ (คู่)", "ฟองน้ำ200 (โหล)", "ฟองน้ำ212 (โหล)", "กล่อง", "กล่องที่"]
    ws_sum.append(headers)
    for cell in ws_sum[1]:
        cell.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C62828")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    g_box = 1
    for i, b in enumerate(all_branches):
        box_nums = ", ".join(f"{g_box+k}/{total_boxes}" for k in range(len(b["boxes"])))
        ws_sum.append([i+1, b["name"], b["name_th"], str(b["upfront"]),
                       b["canvas"], b["foam200"] / 12, b["foam212"] / 12,
                       len(b["boxes"]), box_nums])
        g_box += len(b["boxes"])

    for col in ws_sum.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws_sum.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)
    ws_sum.freeze_panes = "A2"

    # Sheet ตามสาขา
    g_box = 1
    for b in all_branches:
        for box_idx, box in enumerate(b["boxes"]):
            box_label = f"{g_box}/{total_boxes}"
            sname = b["name"][:25] if len(b["boxes"]) == 1 else f"{b['name'][:18]} {g_box}-{total_boxes}"
            sname = sname[:31].translate(str.maketrans('', '', '/\\*?[]:\'"'))

            ws = wb_out.create_sheet(sname)
            for merged in tpl_ws.merged_cells.ranges:
                ws.merge_cells(str(merged))
            for col_l, cd in tpl_ws.column_dimensions.items():
                ws.column_dimensions[col_l].width = cd.width
            for row_n, rd in tpl_ws.row_dimensions.items():
                ws.row_dimensions[row_n].height = rd.height

            for row in tpl_ws.iter_rows():
                for cell in row:
                    nc = ws.cell(row=cell.row, column=cell.column)
                    if cell.has_style:
                        nc.font = copy.copy(cell.font)
                        nc.border = copy.copy(cell.border)
                        nc.fill = copy.copy(cell.fill)
                        nc.alignment = copy.copy(cell.alignment)
                    if cell.value and not str(cell.value).startswith("=VLOOKUP"):
                        nc.value = cell.value

            ws["F3"] = b["upfront"] if str(b["upfront"]).isdigit() else ""
            ws["D3"] = b["name"]
            ws["E4"] = b["name_th"] or b["name"]
            ws["D5"] = f"   {company}"
            ws["E6"] = invoice_no
            ws["E7"] = invoice_date if invoice_date else ""
            ws["A9"] = f"กล่องที่  {box_label}        รวม  {total_boxes}  กล่อง        ({box['label']})"
            ws["A9"].font = Font(name="AngsanaUPC", size=25, bold=True)
            ws["A9"].alignment = Alignment(horizontal="center", vertical="center")
            g_box += 1

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()

# ─── Session state ─────────────────────────────────────────────────────────────
if "parsed_files" not in st.session_state:
    st.session_state.parsed_files = {}
if "all_branches" not in st.session_state:
    st.session_state.all_branches = []

# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ ตั้งค่า")
    st.markdown("---")
    st.markdown("### 📋 เงื่อนไขกล่อง")
    st.markdown("""
- 👟 ผ้าใบ **12 คู่** = 1 กล่อง
- 🩴 ฟองน้ำ **2 โหล** = 1 กล่อง
- 🔀 ฟองน้ำ **1 โหล** + ผ้าใบ **6 คู่** = 1 กล่อง
    """)
    st.markdown("---")
    st.markdown("### 📄 ข้อมูลใบรับสินค้า")
    company = st.text_input("บริษัท", value="นันยางมาร์เก็ตติ้ง จำกัด")
    invoice_no = st.text_input("Invoice No.", placeholder="II690423-007")
    invoice_date = st.date_input("วันที่", value=date.today())

# ─── Main ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="brand">
  <h1>📦 ไทวัสดุ — คำนวณกล่องสินค้า</h1>
  <p>อัปโหลดใบแบ่งสินค้า PDF · คำนวณกล่องอัตโนมัติ · ดาวน์โหลด Excel แยกสาขา · ฟรี 100%</p>
</div>
""", unsafe_allow_html=True)

st.subheader("📂 อัปโหลดใบแบ่งสินค้า (PDF)")
uploaded = st.file_uploader("ลากไฟล์มาวางที่นี่ หรือคลิกเพื่อเลือก",
                             type=["pdf"], accept_multiple_files=True,
                             label_visibility="collapsed")

col1, col2 = st.columns([2, 8])
with col1:
    analyze_btn = st.button("🔍 วิเคราะห์", type="primary", disabled=not uploaded)
with col2:
    if st.button("🗑️ ล้างทั้งหมด"):
        st.session_state.parsed_files = {}
        st.session_state.all_branches = []
        st.rerun()

# ─── Analyse ───────────────────────────────────────────────────────────────────
if analyze_btn and uploaded:
    st.session_state.parsed_files = {}
    st.session_state.all_branches = []
    merged_map = {}

    prog = st.progress(0, text="กำลังอ่าน PDF...")
    for i, f in enumerate(uploaded):
        prog.progress((i) / len(uploaded), text=f"กำลังอ่าน: {f.name}")
        parsed = parse_pdf(f.read(), f.name)
        st.session_state.parsed_files[f.name] = parsed

        for code, b in parsed["branches"].items():
            key = str(code)
            if key in merged_map:
                merged_map[key]["canvas"] += b["canvas"]
                merged_map[key]["foam200"] += b["foam200"]
                merged_map[key]["foam212"] += b["foam212"]
                merged_map[key]["srcFiles"].append(f.name)
            else:
                merged_map[key] = {**b, "upfront": code, "srcFiles": [f.name]}
        prog.progress((i+1) / len(uploaded))

    for b in merged_map.values():
        foam_total = (b["foam200"] + b["foam212"]) / 12
        b["boxes"] = calc_boxes(b["canvas"], foam_total)

    st.session_state.all_branches = list(merged_map.values())
    prog.empty()

# ─── Verify PO vs Dist ─────────────────────────────────────────────────────────
if st.session_state.parsed_files:
    with st.expander("🔎 ตรวจสอบยอด ใบสั่งซื้อ vs ใบแบ่งสินค้า", expanded=True):
        for fname, parsed in st.session_state.parsed_files.items():
            st.markdown(f"**📄 {fname}** (PO: {parsed['po_no'] or 'ไม่พบ'})")
            verify = verify_po_vs_dist(parsed["po_items"], parsed["branches"])
            cols = st.columns(3)
            labels = {"canvas": "ผ้าใบ (คู่)", "foam200": "ฟองน้ำ200 (คู่)", "foam212": "ฟองน้ำ212 (คู่)"}
            for ci, (key, lbl) in enumerate(labels.items()):
                v = verify[key]
                icon = "✅" if v["ok"] else "❌"
                diff_txt = "" if v["ok"] else f" (ต่าง {v['diff']:+.0f})"
                cols[ci].metric(f"{icon} {lbl}", f"{v['po']:.0f}", f"แบ่ง: {v['dist']:.0f}{diff_txt}")
            st.divider()

# ─── Results ───────────────────────────────────────────────────────────────────
branches = st.session_state.all_branches
if branches:
    total_boxes = sum(len(b["boxes"]) for b in branches)
    total_pairs = sum(b["canvas"] + b["foam200"] + b["foam212"] for b in branches)

    st.markdown("---")
    st.subheader("📊 สรุปผลการคำนวณ")
    c1, c2, c3, c4 = st.columns(4)
    for col, num, lbl in [
        (c1, len(uploaded) if uploaded else "-", "ไฟล์"),
        (c2, len(branches), "สาขา"),
        (c3, total_boxes, "กล่องทั้งหมด"),
        (c4, f"{total_pairs:,.0f}", "คู่ทั้งหมด"),
    ]:
        col.markdown(f'<div class="stat-box"><div class="num">{num}</div><div class="lbl">{lbl}</div></div>',
                     unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("🏪 รายละเอียดแต่ละสาขา")
    g_box = 1
    for b in branches:
        bx_count = len(b["boxes"])
        box_nums = " ".join(f'<span class="badge b-num">{g_box+k}/{total_boxes}</span>'
                            for k in range(bx_count))
        detail = " ".join(f'<span class="badge b-{bx["type"]}">{bx["label"]}</span>'
                          for bx in b["boxes"])
        st.markdown(f"""
        <div class="branch-card">
          <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:.5rem">
            <div>
              <div style="font-weight:700;font-size:1rem">{b["name"]}</div>
              <div style="font-size:.8rem;color:#888">{b["name_th"]} · รหัส {b["upfront"]}</div>
              <div style="font-size:.85rem;color:#555;margin-top:.3rem">
                {'🎽 ผ้าใบ <b>'+str(b["canvas"])+'</b> คู่&nbsp;&nbsp;' if b["canvas"] else ''}
                {'🩴 ฟองน้ำ200 <b>'+str(b["foam200"])+'</b> คู่&nbsp;&nbsp;' if b["foam200"] else ''}
                {'🩴 ฟองน้ำ212 <b>'+str(b["foam212"])+'</b> คู่' if b["foam212"] else ''}
              </div>
            </div>
            <div style="text-align:right">
              <div style="font-size:.8rem;color:#888">กล่องที่</div>
              <div>{box_nums}</div>
            </div>
          </div>
          <div style="margin-top:.5rem">{detail}</div>
        </div>
        """, unsafe_allow_html=True)
        g_box += bx_count

    st.markdown("---")
    st.subheader("📥 ดาวน์โหลด Excel")
    col_dl, col_info = st.columns([2, 4])
    with col_dl:
        if st.button("⚙️ สร้าง Excel", type="primary"):
            with st.spinner("กำลังสร้างไฟล์..."):
                try:
                    xlsx = generate_excel(branches, company, invoice_no, invoice_date, total_boxes)
                    st.download_button(
                        "📥 ดาวน์โหลด Excel",
                        data=xlsx,
                        file_name=f"ไทวัสดุ_ใบรับสินค้า_{invoice_no or 'export'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
    with col_info:
        st.info(f"จะสร้าง **{total_boxes + 1} sheet** — สรุปทุกสาขา + ใบรับสินค้า {total_boxes} ใบ")
