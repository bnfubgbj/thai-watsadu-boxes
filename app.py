import streamlit as st
import google.generativeai as genai
import base64
import json
import math
import io
import copy
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os
import fitz  # PyMuPDF

def pdf_to_images_b64(pdf_bytes: bytes) -> list[str]:
    """แปลง PDF แต่ละหน้าเป็น base64 PNG"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        mat = fitz.Matrix(2, 2)  # 2x zoom = 144 DPI
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        images.append(base64.standard_b64encode(img_bytes).decode())
    doc.close()
    return images


# ─── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ไทวัสดุ — คำนวณกล่องสินค้า",
    page_icon="📦",
    layout="wide",
)

# ─── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background: #fafafa; }
    .stApp header { background: transparent; }
    div[data-testid="stSidebar"] { background: #1a1a2e; }
    div[data-testid="stSidebar"] * { color: #eee !important; }
    .brand-header {
        background: linear-gradient(135deg, #c62828 0%, #e53935 100%);
        color: white; padding: 1.2rem 1.5rem; border-radius: 12px;
        margin-bottom: 1.5rem; display: flex; align-items: center; gap: 12px;
    }
    .brand-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
    .brand-header p { margin: 0; opacity: .85; font-size: .9rem; }
    .rule-card {
        background: #fff3e0; border-left: 4px solid #ef6c00;
        padding: .75rem 1rem; border-radius: 0 8px 8px 0; margin-bottom: .5rem;
        font-size: .88rem; color: #4e342e;
    }
    .stat-box {
        background: white; border: 1px solid #e0e0e0; border-radius: 10px;
        padding: 1rem; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,.06);
    }
    .stat-box .num { font-size: 2rem; font-weight: 700; color: #c62828; }
    .stat-box .lbl { font-size: .8rem; color: #888; margin-top: .2rem; }
    .branch-card {
        background: white; border: 1px solid #e8e8e8; border-radius: 10px;
        padding: 1rem 1.2rem; margin-bottom: .6rem;
        box-shadow: 0 1px 3px rgba(0,0,0,.05);
    }
    .branch-name { font-weight: 700; font-size: 1rem; color: #1a1a2e; }
    .branch-code { font-size: .8rem; color: #888; }
    .box-badge {
        display: inline-block; padding: 2px 10px; border-radius: 20px;
        font-size: .78rem; font-weight: 600; margin: 2px;
    }
    .badge-mixed { background: #fff8e1; color: #f57f17; border: 1px solid #ffe082; }
    .badge-foam  { background: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; }
    .badge-canvas{ background: #ffebee; color: #c62828; border: 1px solid #ef9a9a; }
    .badge-num   { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }
    .file-done { color: #2e7d32; font-weight: 600; }
    .file-err  { color: #c62828; font-weight: 600; }
    .stProgress > div > div { background-color: #c62828 !important; }
</style>
""", unsafe_allow_html=True)

# ─── Load branch master data ───────────────────────────────────────────────────
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

# ─── Box calculation logic ─────────────────────────────────────────────────────
def calc_boxes(canvas_pairs: int, foam_dozen: float) -> list[dict]:
    """
    เงื่อนไข:
    - ผ้าใบ 12 คู่ = 1 กล่อง
    - ฟองน้ำ 2 โหล = 1 กล่อง
    - ฟองน้ำ 1 โหล + ผ้าใบ 6 คู่ = 1 กล่อง (กล่องผสม)
    """
    boxes = []
    cp = canvas_pairs
    fd = foam_dozen

    # กล่องผสมก่อน (ประหยัดสุด)
    mixed = min(int(fd), cp // 6)
    for _ in range(mixed):
        boxes.append({"type": "mixed", "label": "ฟองน้ำ 1 โหล + ผ้าใบ 6 คู่"})
        fd -= 1
        cp -= 6

    # ฟองน้ำล้วน 2 โหล/กล่อง
    while fd >= 2:
        boxes.append({"type": "foam", "label": "ฟองน้ำ 2 โหล"})
        fd -= 2

    # ฟองน้ำเศษ 1 โหล
    if fd >= 1:
        if cp >= 6:
            boxes.append({"type": "mixed", "label": "ฟองน้ำ 1 โหล + ผ้าใบ 6 คู่"})
            cp -= 6
        else:
            boxes.append({"type": "foam", "label": "ฟองน้ำ 1 โหล"})
        fd = 0

    # ผ้าใบล้วน
    while cp >= 12:
        boxes.append({"type": "canvas", "label": "ผ้าใบ 12 คู่"})
        cp -= 12
    if cp > 0:
        boxes.append({"type": "canvas", "label": f"ผ้าใบ {cp} คู่"})

    return boxes

# ─── AI: analyse one file ──────────────────────────────────────────────────────
def analyse_file(file_bytes: bytes, mime: str, api_key: str) -> list[dict]:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    prompt = """คุณคือผู้เชี่ยวชาญอ่านใบแบ่งสินค้ารองเท้าของไทวัสดุ (CRC Thai Watsadu)

จากรูปภาพนี้ ให้ดึงข้อมูลทุกสาขาออกมา แยกประเภทสินค้าเป็น:
1. รองเท้าผ้าใบ (Canvas) — หน่วย: คู่ (pair)
2. รองเท้าฟองน้ำ 200 — หน่วย: โหล (dozen)
3. รองเท้าฟองน้ำ 212 — หน่วย: โหล (dozen)

ตอบเป็น JSON เท่านั้น ไม่มีคำอธิบาย ไม่มี backtick:
{"branches":[{"upfront":"รหัส 5 หลัก","name":"ชื่อ EN","nameTH":"ชื่อ TH","canvasPairs":0,"foam200Dozen":0,"foam212Dozen":0}],"invoiceNo":""}

หมายเหตุ: EACH = คู่, DOZEN = โหล (12 คู่)"""

    if mime == "application/pdf":
        pages_b64 = pdf_to_images_b64(file_bytes)
        all_parsed_branches = []
        invoice_no_found = ""
        for page_b64 in pages_b64:
            img_data = base64.b64decode(page_b64)
            resp = model.generate_content([
                {"mime_type": "image/png", "data": img_data},
                prompt,
            ])
            raw = resp.text.strip()
            try:
                page_parsed = json.loads(raw.replace("```json", "").replace("```", "").strip())
                all_parsed_branches.extend(page_parsed.get("branches", []))
                if not invoice_no_found:
                    invoice_no_found = page_parsed.get("invoiceNo", "")
            except Exception:
                pass
        parsed = {"branches": all_parsed_branches, "invoiceNo": invoice_no_found}
    else:
        from PIL import Image
        img = Image.open(io.BytesIO(file_bytes))
        resp = model.generate_content([img, prompt])
        raw = resp.text.strip()
        parsed = json.loads(raw.replace("```json", "").replace("```", "").strip())

    result = []
    for b in parsed.get("branches", []):
        foam = (b.get("foam200Dozen") or 0) + (b.get("foam212Dozen") or 0)
        cp = b.get("canvasPairs") or 0
        upfront_str = str(b.get("upfront", "")).strip()
        upfront_int = int(upfront_str) if upfront_str.isdigit() else 0

        # lookup ชื่อจาก master
        master = BRANCH_MASTER.get(upfront_int, {})
        result.append({
            "upfront": upfront_int or upfront_str,
            "name": master.get("name_en") or b.get("name", ""),
            "nameTH": master.get("name_th") or b.get("nameTH", ""),
            "canvasPairs": cp,
            "foamDozen": foam,
            "boxes": calc_boxes(cp, foam),
        })
    return result, parsed.get("invoiceNo", "")

# ─── Excel generation ──────────────────────────────────────────────────────────
def generate_excel(branches: list[dict], company: str, invoice_no: str,
                   invoice_date, total_boxes: int) -> bytes:
    wb_tpl = load_workbook(TEMPLATE_PATH)
    tpl_ws = wb_tpl["ใบปะหน้า"]

    from openpyxl import Workbook
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    # ── Sheet สรุปทุกสาขา ──
    ws_sum = wb_out.create_sheet("สรุปทุกสาขา")
    headers = ["ลำดับ", "สาขา (EN)", "สาขา (TH)", "รหัส", "ผ้าใบ (คู่)", "ฟองน้ำ (โหล)", "กล่อง", "กล่องที่"]
    ws_sum.append(headers)
    for cell in ws_sum[1]:
        cell.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C62828")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    g_box = 1
    for i, b in enumerate(branches):
        box_nums = ", ".join(f"{g_box+k}/{total_boxes}" for k in range(len(b["boxes"])))
        ws_sum.append([
            i + 1,
            b["name"],
            b["nameTH"],
            str(b["upfront"]),
            b["canvasPairs"],
            b["foamDozen"],
            len(b["boxes"]),
            box_nums,
        ])
        g_box += len(b["boxes"])

    for col in ws_sum.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws_sum.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)

    ws_sum.freeze_panes = "A2"

    # ── Sheet ต่อไปตามสาขา (copy template) ──
    g_box = 1
    for b in branches:
        branch_boxes = b["boxes"]
        for box_idx, box in enumerate(branch_boxes):
            box_num = g_box + box_idx
            box_label = f"{box_num}/{total_boxes}"

            sheet_name = b["name"][:25] if len(branch_boxes) == 1 else f"{b['name'][:20]} {box_num}-{total_boxes}"
            sheet_name = sheet_name[:31].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "").replace(":", "")

            ws = wb_out.create_sheet(sheet_name)

            # copy column widths & row heights from template
            for col_letter, cd in tpl_ws.column_dimensions.items():
                ws.column_dimensions[col_letter].width = cd.width
            for row_num, rd in tpl_ws.row_dimensions.items():
                ws.row_dimensions[row_num].height = rd.height

            # copy merge cells
            for merged in tpl_ws.merged_cells.ranges:
                ws.merge_cells(str(merged))

            # copy cells
            for row in tpl_ws.iter_rows():
                for cell in row:
                    new_cell = ws.cell(row=cell.row, column=cell.column)
                    if cell.has_style:
                        new_cell.font = copy.copy(cell.font)
                        new_cell.border = copy.copy(cell.border)
                        new_cell.fill = copy.copy(cell.fill)
                        new_cell.number_format = cell.number_format
                        new_cell.protection = copy.copy(cell.protection)
                        new_cell.alignment = copy.copy(cell.alignment)
                    # copy value แต่ข้าม VLOOKUP formula → แทนด้วยค่าจริง
                    if cell.value and str(cell.value).startswith("=VLOOKUP"):
                        pass  # จะกรอกด้านล่าง
                    else:
                        new_cell.value = cell.value

            # กรอกข้อมูลจริง
            upfront_val = b["upfront"] if str(b["upfront"]).isdigit() else ""
            ws["F3"] = int(upfront_val) if str(upfront_val).isdigit() else upfront_val
            ws["D3"] = b["name"]          # ชื่อ EN แทน VLOOKUP
            ws["E4"] = b["nameTH"] or b["name"]   # ชื่อ TH
            ws["D5"] = f"   {company}"
            ws["E6"] = invoice_no
            ws["E7"] = invoice_date if invoice_date else ""

            # กล่องที่ X/Y รวม Y กล่อง
            detail = box["label"]
            ws["A9"] = f"กล่องที่  {box_label}        รวม  {total_boxes}  กล่อง        ({detail})"
            ws["A9"].font = Font(name="AngsanaUPC", size=25, bold=True)
            ws["A9"].alignment = Alignment(horizontal="center", vertical="center")

        g_box += len(branch_boxes)

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()

# ─── Session state ─────────────────────────────────────────────────────────────
if "all_branches" not in st.session_state:
    st.session_state.all_branches = []
if "file_results" not in st.session_state:
    st.session_state.file_results = {}

# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ ตั้งค่า")
    # ดึง key จาก Streamlit Secrets ก่อน ถ้าไม่มีค่อยให้ user กรอก
    default_key = st.secrets.get("GEMINI_API_KEY", "")
    api_key = st.text_input("Google Gemini API Key", type="password",
                            value=default_key,
                            placeholder="AIza...",
                            help="ดูได้จาก aistudio.google.com/app/apikey")
    st.markdown("---")
    st.markdown("### 📋 เงื่อนไขการบรรจุกล่อง")
    st.markdown("""
- 👟 **ผ้าใบ** 12 คู่ = 1 กล่อง
- 🩴 **ฟองน้ำ** 2 โหล = 1 กล่อง
- 🔀 **ผสม** ฟองน้ำ 1 โหล + ผ้าใบ 6 คู่ = 1 กล่อง
- ฟองน้ำมี 2 ชนิด: **200** และ **212**
    """)
    st.markdown("---")
    st.markdown("### 📄 ข้อมูลใบรับสินค้า")
    company = st.text_input("บริษัท / ผู้จัดส่ง", placeholder="นันยางมาร์เก็ตติ้ง จำกัด")
    invoice_no = st.text_input("Invoice No.", placeholder="II690423-007")
    invoice_date = st.date_input("วันที่", value=date.today())

# ─── Main ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="brand-header">
  <div>
    <h1>📦 ไทวัสดุ — คำนวณกล่องสินค้า</h1>
    <p>อัปโหลดใบแบ่งสินค้า PDF หลายไฟล์ · AI อ่านอัตโนมัติ · ดาวน์โหลด Excel แยกสาขา</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── Upload zone ───────────────────────────────────────────────────────────────
st.subheader("📂 อัปโหลดใบแบ่งสินค้า")
uploaded = st.file_uploader(
    "ลากไฟล์มาวางที่นี่ หรือคลิกเพื่อเลือก",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

col_btn1, col_btn2, col_btn3 = st.columns([2, 2, 6])
with col_btn1:
    analyze_btn = st.button("🔍 วิเคราะห์ทั้งหมด", type="primary",
                            disabled=not (uploaded and api_key))
with col_btn2:
    if st.button("🗑️ ล้างผลลัพธ์"):
        st.session_state.all_branches = []
        st.session_state.file_results = {}
        st.rerun()

if not api_key:
    st.info("💡 กรอก **Google Gemini API Key** ในแถบซ้ายก่อนเริ่มใช้งาน")

# ─── Analyse ───────────────────────────────────────────────────────────────────
if analyze_btn and uploaded and api_key:
    st.session_state.all_branches = []
    st.session_state.file_results = {}

    progress = st.progress(0, text="เริ่มวิเคราะห์...")
    status_area = st.empty()

    merged_map = {}

    for i, f in enumerate(uploaded):
        progress.progress((i) / len(uploaded), text=f"กำลังอ่าน: {f.name}")
        status_area.info(f"⏳ วิเคราะห์ไฟล์ {i+1}/{len(uploaded)}: **{f.name}**")

        try:
            mime = f.type if f.type else "image/jpeg"
            if f.name.lower().endswith(".pdf"):
                mime = "application/pdf"

            branches, inv = analyse_file(f.read(), mime, api_key)

            if inv and not invoice_no:
                invoice_no = inv

            st.session_state.file_results[f.name] = {"status": "done", "branches": branches}

            # merge สาขาซ้ำกัน
            for b in branches:
                key = str(b["upfront"])
                if key in merged_map:
                    merged_map[key]["canvasPairs"] += b["canvasPairs"]
                    merged_map[key]["foamDozen"] += b["foamDozen"]
                    merged_map[key]["srcFiles"].append(f.name)
                else:
                    merged_map[key] = {**b, "srcFiles": [f.name]}

        except Exception as e:
            st.session_state.file_results[f.name] = {"status": "error", "error": str(e)}

        progress.progress((i + 1) / len(uploaded), text=f"เสร็จ {i+1}/{len(uploaded)}")

    # คำนวณกล่องใหม่หลัง merge
    for b in merged_map.values():
        b["boxes"] = calc_boxes(b["canvasPairs"], b["foamDozen"])

    st.session_state.all_branches = list(merged_map.values())
    status_area.success(f"✅ วิเคราะห์เสร็จ — พบ **{len(st.session_state.all_branches)} สาขา** จาก {len(uploaded)} ไฟล์")
    progress.empty()

# ─── Show file status ──────────────────────────────────────────────────────────
if uploaded and st.session_state.file_results:
    with st.expander("📋 สถานะไฟล์", expanded=False):
        for f in uploaded:
            res = st.session_state.file_results.get(f.name)
            if not res:
                st.write(f"⏳ {f.name} — รอ")
            elif res["status"] == "done":
                n = len(res["branches"])
                st.markdown(f'<span class="file-done">✓ {f.name} — พบ {n} สาขา</span>', unsafe_allow_html=True)
            else:
                st.markdown(f'<span class="file-err">✗ {f.name} — {res.get("error","ผิดพลาด")}</span>', unsafe_allow_html=True)

# ─── Results ───────────────────────────────────────────────────────────────────
branches = st.session_state.all_branches
if branches:
    total_boxes = sum(len(b["boxes"]) for b in branches)
    total_pairs = sum(b["canvasPairs"] + b["foamDozen"] * 12 for b in branches)

    st.markdown("---")
    st.subheader("📊 สรุปผลการคำนวณ")

    c1, c2, c3, c4 = st.columns(4)
    for col, num, lbl in [
        (c1, len(uploaded) if uploaded else "-", "ไฟล์ที่วิเคราะห์"),
        (c2, len(branches), "สาขา"),
        (c3, total_boxes, "กล่องทั้งหมด"),
        (c4, f"{total_pairs:,}", "คู่ทั้งหมด"),
    ]:
        col.markdown(f"""
        <div class="stat-box">
            <div class="num">{num}</div>
            <div class="lbl">{lbl}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("🏪 รายละเอียดแต่ละสาขา")

    g_box = 1
    for b in branches:
        bx_count = len(b["boxes"])
        box_nums_html = " ".join(
            f'<span class="box-badge badge-num">{g_box+k}/{total_boxes}</span>'
            for k in range(bx_count)
        )
        detail_html = " ".join(
            f'<span class="box-badge badge-{bx["type"]}">{bx["label"]}</span>'
            for bx in b["boxes"]
        )
        st.markdown(f"""
        <div class="branch-card">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:.5rem">
                <div>
                    <div class="branch-name">{b["name"]}</div>
                    <div class="branch-code">{b["nameTH"]} &nbsp;·&nbsp; รหัส {b["upfront"]}</div>
                    <div style="margin-top:.4rem;font-size:.85rem;color:#555">
                        {'🎽 ผ้าใบ <b>'+str(b["canvasPairs"])+'</b> คู่' if b["canvasPairs"] else ''}
                        {'&nbsp;&nbsp;🩴 ฟองน้ำ <b>'+str(b["foamDozen"])+'</b> โหล' if b["foamDozen"] else ''}
                    </div>
                </div>
                <div style="text-align:right">
                    <div style="font-size:.82rem;color:#888;margin-bottom:.3rem">กล่องที่</div>
                    <div>{box_nums_html}</div>
                </div>
            </div>
            <div style="margin-top:.5rem">{detail_html}</div>
        </div>
        """, unsafe_allow_html=True)
        g_box += bx_count

    # ─── Download Excel ────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📥 ดาวน์โหลด Excel")

    col_dl, col_info = st.columns([2, 4])
    with col_dl:
        if st.button("⚙️ สร้าง Excel", type="primary"):
            with st.spinner("กำลังสร้างไฟล์..."):
                try:
                    xlsx_bytes = generate_excel(
                        branches, company, invoice_no, invoice_date, total_boxes
                    )
                    st.download_button(
                        label="📥 ดาวน์โหลด Excel",
                        data=xlsx_bytes,
                        file_name=f"ไทวัสดุ_ใบรับสินค้า_{invoice_no or 'export'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
    with col_info:
        st.info(f"จะสร้าง **{total_boxes + 1} sheet** — sheet สรุปทุกสาขา + sheet ใบรับสินค้าแยกตามกล่อง ({total_boxes} ใบ)")
