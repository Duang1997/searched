import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
import datetime

# --- ตั้งค่าหน้าเพจ ---
st.set_page_config(page_title="ระบบบันทึกการตรวจค้น/ตรวจยึด (CIB)", layout="wide") 

# --- ตกแต่ง UI ธีม CIB ---
st.markdown("""
<style>
    .cib-header { background-color: #00204a; padding: 15px; border-radius: 5px; color: #f9bc0f; text-align: center; font-family: sans-serif; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.3); }
    .cib-header h1 { color: #f9bc0f; margin: 0; font-size: 28px; font-weight: bold; }
    .cib-header p { color: #ffffff; margin: 0; font-size: 16px; }
</style>
<div class="cib-header">
    <h1>ตำรวจสอบสวนกลาง (Central Investigation Bureau)</h1>
    <p>ระบบจัดทำบันทึกการตรวจค้น / ตรวจยึด</p>
</div>
""", unsafe_allow_html=True)

# --- ฟังก์ชันจัดการวันที่ (แปลงเป็น A.D. รูปแบบ dd/mm/yyyy) ---
def format_ad_date(date_obj):
    if not date_obj: return ""
    return date_obj.strftime("%d/%m/%Y")

# --- จัดการ Session State ---
if "seized_df" not in st.session_state:
    st.session_state.seized_df = pd.DataFrame([{"ลำดับ": "1", "รายการสิ่งของ": "", "จำนวน": "", "หมายเหตุ": ""}])
if "search_officer_df" not in st.session_state:
    st.session_state.search_officer_df = pd.DataFrame([{"ยศ": "พ.ต.ต.", "ชื่อ-นามสกุล": "สุวิจักขณ์ รัตนพันธ์", "ตำแหน่ง": "สว.กก.๓ บก.ป."}])

# ==========================================
# ส่วนที่ 1: ข้อมูลวันเวลาและสถานที่
# ==========================================
st.header("ส่วนที่ 1: ข้อมูลวันเวลาและสถานที่")
col1, col2 = st.columns(2)
with col1:
    record_location = st.text_input("สถานที่บันทึก", placeholder="เช่น กองกำกับการ 3 กองบังคับการปราบปราม")
    record_date = st.date_input("วันที่บันทึก", key="rec_date")
    record_time = st.text_input("เวลาที่บันทึก", placeholder="เช่น 08.30")
with col2:
    search_location = st.text_input("สถานที่ตรวจค้น", placeholder="เช่น ร้านโอเครถเช่า...")
    search_date = st.date_input("วันที่ตรวจค้น/ตรวจยึด", key="search_date")
    search_time = st.text_input("เวลาที่ตรวจค้น/ตรวจยึด", placeholder="เช่น 06.00")
    search_end_time = st.text_input("เวลาที่เสร็จสิ้นการตรวจค้น", placeholder="เช่น 08.00")

st.divider()

# ==========================================
# ส่วนที่ 2: อำนาจในการตรวจค้นและผู้นำตรวจค้น
# ==========================================
st.header("ส่วนที่ 2: อำนาจในการตรวจค้นและผู้นำตรวจค้น")
col3, col4 = st.columns(2)
with col3:
    st.subheader("รายละเอียดหมายค้น")
    warrant_court = st.text_input("ศาลที่ออกหมายค้น", placeholder="เช่น ศาลจังหวัดอุบลราชธานี")
    warrant_no = st.text_input("หมายค้นที่", placeholder="เช่น 47/2568")
    warrant_date = st.date_input("ลงวันที่ (หมายค้น)", key="warrant_date")
with col4:
    st.subheader("เจ้าบ้าน/ผู้นำตรวจค้น")
    leader_name_1 = st.text_input("ชื่อ-นามสกุล ผู้นำตรวจค้น (1)", placeholder="ระบุชื่อ...")
    leader_status_1 = st.text_input("สถานะ (1)", placeholder="เช่น ผู้ดูแลสถานที่, เจ้าบ้าน")
    leader_name_2 = st.text_input("ชื่อ-นามสกุล ผู้นำตรวจค้น (2)", placeholder="ระบุชื่อ (ถ้ามี)...")

st.divider()

# ==========================================
# ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ
# ==========================================
st.header("ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ")
commanders = st.text_area("ภายใต้การอำนวยการสั่งการของ", placeholder="พล.ต.ท.จิรภพ ภูริเดช ผบช.ก....")

st.subheader("เจ้าหน้าที่ผู้ทำการตรวจค้น")
off_mode = st.radio("รูปแบบการเพิ่มรายชื่อ", ["กรอกผ่านตารางในเว็บ", "อัปโหลดไฟล์ Excel"], horizontal=True)

if off_mode == "กรอกผ่านตารางในเว็บ":
    edited_officers = st.data_editor(st.session_state.search_officer_df, num_rows="dynamic", use_container_width=True)
    valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].astype(str).str.strip() != ""]
else:
    st.info("💡 รูปแบบหัวตาราง Excel ที่ระบบต้องการ: `ยศ` | `ชื่อ-นามสกุล` | `ตำแหน่ง`")
    uploaded_off = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])
    if uploaded_off:
        df_off = pd.read_excel(uploaded_off, dtype=str)
        df_off.columns = df_off.columns.str.strip()
        if "ชื่อ-นามสกุล" in df_off.columns:
            valid_officers = df_off[df_off["ชื่อ-นามสกุล"].notna() & (df_off["ชื่อ-นามสกุล"].astype(str).str.strip() != "")]
        else:
            valid_officers = pd.DataFrame(columns=["ยศ", "ชื่อ-นามสกุล", "ตำแหน่ง"])
            st.error("⚠️ ไม่พบคอลัมน์ 'ชื่อ-นามสกุล' ในไฟล์ Excel")
    else:
        valid_officers = pd.DataFrame(columns=["ยศ", "ชื่อ-นามสกุล", "ตำแหน่ง"])

officers_list = []
officer_displays = []
for _, r in valid_officers.iterrows():
    rank = str(r.get('ยศ', '')).replace("nan", "").strip()
    name = str(r.get('ชื่อ-นามสกุล', '')).replace("nan", "").strip()
    pos = str(r.get('ตำแหน่ง', '')).replace("nan", "").strip()
    
    display = f"{rank}{name} {pos}".strip()
    officers_list.append({"rank": rank, "name_only": name, "display": display})
    officer_displays.append(display)

police_team_text = ", ".join(officer_displays)

signature_rows = []
for j in range(0, len(officers_list), 2):
    o1 = officers_list[j]
    o2 = officers_list[j+1] if j+1 < len(officers_list) else {"rank": "", "name_only": ""}
    signature_rows.append({
        "officer1_rank": o1["rank"], "officer1_name": o1["name_only"], 
        "officer2_rank": o2["rank"], "officer2_name": o2["name_only"]
    })

st.divider()

# ==========================================
# ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น
# ==========================================
st.header("ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น")
search_circumstances = st.text_area("พฤติการณ์ในการตรวจค้น/ตรวจยึด", height=150, placeholder="ตามวันเวลาที่แจ้ง เจ้าพนักงานตำรวจชุดตรวจค้น...")
investigator_name = st.text_input("พนักงานสอบสวนผู้รับผิดชอบ (พงส.)", placeholder="เช่น ร.ต.อ. ...")

st.subheader("รายการสิ่งของตรวจยึด")
edited_seized = st.data_editor(st.session_state.seized_df, num_rows="dynamic", use_container_width=True)
seized_items_dict = edited_seized.to_dict('records')

st.divider()

# ==========================================
# ส่วนที่ 5: ภาพประกอบการตรวจค้น
# ==========================================
st.header("ส่วนที่ 5: ภาพประกอบการตรวจค้น")
st.caption("อัปโหลดภาพถ่ายการปฏิบัติงานเพื่อแนบท้ายบันทึก")
col_img1, col_img2 = st.columns(2)
with col_img1:
    img_1 = st.file_uploader("ภาพประกอบที่ 1", type=['png', 'jpg', 'jpeg'])
with col_img2:
    img_2 = st.file_uploader("ภาพประกอบที่ 2", type=['png', 'jpg', 'jpeg'])

st.divider()

# ==========================================
# การประมวลผลและสร้างเอกสาร
# ==========================================
if st.button("💾 สร้างและดาวน์โหลด บันทึกการตรวจค้น/ตรวจยึด", type="primary", use_container_width=True):
    try:
        doc = DocxTemplate("template_search_seizure.docx")
        
        context = {
            "record_location": record_location,
            "record_date_ad": format_ad_date(record_date),
            "record_time": record_time,
            "search_location": search_location,
            "search_date_ad": format_ad_date(search_date),
            "search_time": search_time,
            "search_end_time": search_end_time,
            "warrant_court": warrant_court,
            "warrant_no": warrant_no,
            "warrant_date_ad": format_ad_date(warrant_date),
            "leader_name_1": leader_name_1,
            "leader_status_1": leader_status_1,
            "leader_name_2": leader_name_2,
            "commanders": commanders,
            "police_team_text": police_team_text,
            "signature_rows": signature_rows,
            "search_circumstances": search_circumstances,
            "investigator_name": investigator_name,
            "seized_items": seized_items_dict,
            "img_1": InlineImage(doc, img_1, width=Mm(75)) if img_1 else "",
            "img_2": InlineImage(doc, img_2, width=Mm(75)) if img_2 else ""
        }
        
        doc.render(context)
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        
        st.success("✅ ประมวลผลเอกสารเสร็จสิ้น")
        st.download_button(
            label="⬇️ โหลดไฟล์ บันทึกตรวจค้นตรวจยึด.docx",
            data=bio.getvalue(),
            file_name=f"บันทึกตรวจค้น_{datetime.datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการสร้างเอกสาร: {e}\n(กรุณาตรวจสอบว่ามีไฟล์ template_search_seizure.docx อยู่ในโฟลเดอร์เดียวกับโปรแกรม)")
