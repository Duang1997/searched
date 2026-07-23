import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
import datetime
import os

# --- ตั้งค่าหน้าเพจ ---
st.set_page_config(page_title="ระบบบันทึกการตรวจค้น/ตรวจยึด (CIB)") 

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

# --- ฟังก์ชันจัดการวันที่ (พ.ศ.) ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def format_thai_date(date_obj):
    if not date_obj: return ""
    return f"วันที่ {date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

def format_thai_date_only(date_obj):
    if not date_obj: return ""
    return f"{date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

# --- จัดการ Session State ---
if 'unit_count' not in st.session_state:
    st.session_state.unit_count = 1
if 'leader_count' not in st.session_state:
    st.session_state.leader_count = 1

def add_unit():
    st.session_state.unit_count += 1

def add_leader():
    st.session_state.leader_count += 1

# ==========================================
# ส่วนที่ 1: ข้อมูลวันเวลาและสถานที่
# ==========================================
st.header("ส่วนที่ 1: ข้อมูลวันเวลาและสถานที่")
record_location = st.text_input("สถานที่บันทึก", placeholder="เช่น กองกำกับการ 3 กองบังคับการปราบปราม")
record_date = st.date_input("วันที่บันทึก", key="rec_date")
record_time = st.text_input("เวลาที่บันทึก", placeholder="เช่น 08.30")
search_location = st.text_input("สถานที่ตรวจค้น", placeholder="เช่น ร้านโอเครถเช่า...")
search_date = st.date_input("วันที่ตรวจค้น/ตรวจยึด", key="search_date")
search_time = st.text_input("เวลาที่ตรวจค้น/ตรวจยึด", placeholder="เช่น 06.00")
search_end_time = st.text_input("เวลาที่เสร็จสิ้นการตรวจค้น", placeholder="เช่น 08.00")
st.divider()

# ==========================================
# ส่วนที่ 2: อำนาจในการตรวจค้นและผู้นำตรวจค้น
# ==========================================
st.header("ส่วนที่ 2: อำนาจในการตรวจค้นและผู้นำตรวจค้น")
warrant_court = st.text_input("ศาลที่ออกหมายค้น", placeholder="เช่น ศาลอาญา")
warrant_no = st.text_input("หมายค้นที่", placeholder="เช่น 3/26")
warrant_date = st.date_input("ลงวันที่ (หมายค้น)", key="warrant_date")

warrant_text = f"หมายค้นของ {warrant_court} ที่ {warrant_no} ลงวันที่ {format_thai_date_only(warrant_date)}"

st.subheader("ผู้นำตรวจค้น")
leaders = []
for i in range(st.session_state.leader_count):
    with st.container(border=True):
        lname = st.text_input(f"ชื่อ-นามสกุล ผู้นำตรวจค้นคนที่ {i+1}", key=f"l_name_{i}")
        lstatus = st.text_input(f"เกี่ยวข้องเป็น (สถานะ) คนที่ {i+1}", placeholder="เช่น ผู้ดูแลสถานที่, เจ้าบ้าน", key=f"l_status_{i}")
        if lname.strip():
            leaders.append({"name": lname.strip(), "status": lstatus.strip()})
st.button("➕ เพิ่มผู้นำตรวจค้น", on_click=add_leader)

# ประมวลผลข้อความผู้นำตรวจค้น (เฉพาะชื่อและสถานะ)
leader_texts = [f"{l['name']} เกี่ยวข้องเป็น {l['status']}" for l in leaders]
leaders_intro_text = " และ ".join(leader_texts) if leader_texts else ""
st.divider()

# ==========================================
# ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ
# ==========================================
st.header("ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ")
default_cmd = "พล.ต.ต.ณัฐศักดิ์ เชาวนาศัย ผบช.ก., พ.ต.ต.พัฒนศักดิ์ บุบผาสุวรรณ ผบก.ป., พ.ต.อ.สุเทพ โตอิ้ม รอง ผบก.ป., พ.ต.อ.สุริยศักดิ์ จิราวัสน์ ผกก.3 บก.ป., พ.ต.ท.พงษ์พิทักษ์ เหล็กชูชาติ, พ.ต.ท.รัฐมนตรี พันชูกลาง, พ.ต.ท.ณัฐดนัย สีแข่ไตร, พ.ต.ท.ศิษฏ์ พูลวงศ์, พ.ต.ท.พัฒษพงศ์ เสณีแสนเสนา รอง ผกก.3 บก.ป."
default_officer_row = {"ยศ": "พ.ต.ต.", "ชื่อ-นามสกุล": "สุวิจักขณ์ รัตนพันธ์", "ตำแหน่ง": "สว.กก.๓ บก.ป."}

units_data = []
all_officer_displays = []

for i in range(st.session_state.unit_count):
    with st.container(border=True):
        st.subheader(f"🏢 หน่วยตรวจค้นที่ {i+1}")
        unit_name = st.text_input(f"ชื่อหน่วยงาน", value="กก.๓ บก.ป." if i==0 else "", placeholder="เช่น กก.๓ บก.ป.", key=f"u_name_{i}")
        commanders_text = st.text_area(f"ภายใต้อำนวยการสั่งการของ", value=default_cmd if i==0 else "", key=f"cmd_{i}")
        
        df_key = f"officer_df_{i}"
        file_key = f"uploaded_file_id_{i}"
        
        if df_key not in st.session_state:
            st.session_state[df_key] = pd.DataFrame([default_officer_row]) if i==0 else pd.DataFrame([{"ยศ": "พ.ต.ต.", "ชื่อ-นามสกุล": "", "ตำแหน่ง": ""}])
        if file_key not in st.session_state:
            st.session_state[file_key] = None

        up_file = st.file_uploader(f"อัปโหลดไฟล์ Excel (.xlsx) หน่วยที่ {i+1} (ไม่บังคับ)", type=["xlsx"], key=f"up_{i}")
        
        # ตรวจสอบว่ามีการอัปโหลดไฟล์ใหม่หรือไม่
        if up_file is not None:
            if st.session_state[file_key] != up_file.file_id:
                df = pd.read_excel(up_file, dtype=str)
                df.columns = df.columns.str.strip()
                if "ชื่อ-นามสกุล" in df.columns:
                    if "ยศ" not in df.columns: df["ยศ"] = ""
                    if "ตำแหน่ง" not in df.columns: df["ตำแหน่ง"] = ""
                    df = df[["ยศ", "ชื่อ-นามสกุล", "ตำแหน่ง"]].fillna("")
                    st.session_state[df_key] = df
                    st.session_state[file_key] = up_file.file_id
                    st.success(f"✅ โหลดข้อมูลจาก Excel สำเร็จ สามารถแก้ไขในตารางด้านล่างได้ทันที")
                else:
                    st.error("⚠️ ไม่พบคอลัมน์ 'ชื่อ-นามสกุล' ในไฟล์")
                    st.session_state[file_key] = up_file.file_id

        # แสดงตารางให้ผู้ใช้กรอก/แก้ไขข้อมูล
        edited_officers = st.data_editor(st.session_state[df_key], num_rows="dynamic", use_container_width=True, key=f"edit_{i}")
        st.session_state[df_key] = edited_officers
        
        valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].astype(str).str.strip() != ""]

        officers_list = []
        for _, r in valid_officers.iterrows():
            rank = str(r.get('ยศ', '')).replace("nan", "").strip()
            name = str(r.get('ชื่อ-นามสกุล', '')).replace("nan", "").strip()
            pos = str(r.get('ตำแหน่ง', '')).replace("nan", "").strip()
            
            if name:
                display = f"{rank}{name} {pos}".strip()
                officers_list.append({"rank": rank, "name_only": name, "display": display})
                all_officer_displays.append(display)

        signature_rows = []
        for j in range(0, len(officers_list), 2):
            o1 = officers_list[j]
            o2 = officers_list[j+1] if j+1 < len(officers_list) else {"rank": "", "name_only": ""}
            signature_rows.append({
                "officer1_rank": o1["rank"], "officer1_name": o1["name_only"], 
                "officer2_rank": o2["rank"], "officer2_name": o2["name_only"]
            })

        units_data.append({
            "unit_name": unit_name,
            "commanders_text": commanders_text,
            "police_team_text": f"เจ้าพนักงานตำรวจ ({unit_name}) ประกอบด้วย " + ", ".join([o["display"] for o in officers_list]) if officers_list else "",
            "signature_rows": signature_rows
        })

st.button("➕ เพิ่มหน่วยตรวจค้นอื่น", on_click=add_unit)
st.divider()

# ==========================================
# ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น
# ==========================================
st.header("ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น")
search_circumstances = st.text_area("พฤติการณ์ในการตรวจค้น/ตรวจยึด", height=150)
seized_count = st.number_input("จำนวนรายการสิ่งของตรวจยึด (รายการ)", min_value=0, value=1)

handover_opt = st.radio("นำทรัพย์ทั้งหมดส่งมอบให้ใคร", ["พนักงานสอบสวนผู้รับผิดชอบ", "อื่นๆ (ระบุ)"])
if handover_opt == "อื่นๆ (ระบุ)":
    investigator_name = st.text_input("ระบุชื่อ/ตำแหน่งผู้รับมอบ", placeholder="เช่น ร.ต.อ. ...")
else:
    investigator_name = "พนักงานสอบสวนผู้รับผิดชอบ"
st.divider()

# ==========================================
# ส่วนที่ 5: ภาพประกอบการตรวจค้น
# ==========================================
st.header("ส่วนที่ 5: ภาพประกอบการตรวจค้น")
st.caption("อัปโหลดภาพถ่ายการปฏิบัติงานเพื่อแนบท้ายบันทึก")
img_1 = st.file_uploader("ภาพประกอบที่ 1", type=['png', 'jpg', 'jpeg'])
img_2 = st.file_uploader("ภาพประกอบที่ 2", type=['png', 'jpg', 'jpeg'])

st.subheader("ผู้ลงนามรับรองภาพถ่าย")
leader_names = [l["name"] for l in leaders] if 'leaders' in locals() else []
sign_choices = leader_names + all_officer_displays + ["อื่นๆ (ระบุ)"]

img_sign_opt = st.selectbox("เลือกผู้ลงนามภาพถ่าย", sign_choices)
if img_sign_opt == "อื่นๆ (ระบุ)":
    img_signer = st.text_input("ระบุชื่อผู้ลงนามภาพถ่าย")
else:
    img_signer = img_sign_opt
st.divider()

# ==========================================
# การประมวลผลและสร้างเอกสาร
# ==========================================
st.header("ส่วนที่ 6: ดาวน์โหลดเอกสารและไฟล์แนบ")

col_ex, col_pp = st.columns(2)
with col_ex:
    if os.path.exists("บัญชีของกลาง.xlsx"):
        with open("บัญชีของกลาง.xlsx", "rb") as f:
            st.download_button("📊 โหลดแม่แบบ บัญชีตรวจยึด (Excel)", f, file_name="บัญชีของกลาง.xlsx", use_container_width=True)
    else:
        st.button("📊 โหลดแม่แบบ บัญชีตรวจยึด (ไม่พบไฟล์)", disabled=True, use_container_width=True)

with col_pp:
    if os.path.exists("ป้ายของกลาง.pptx"):
        with open("ป้ายของกลาง.pptx", "rb") as f:
            st.download_button("🖼️ โหลดแม่แบบ ป้ายของกลาง (PPTX)", f, file_name="ป้ายของกลาง.pptx", use_container_width=True)
    else:
        st.button("🖼️ โหลดแม่แบบ ป้ายของกลาง (ไม่พบไฟล์)", disabled=True, use_container_width=True)

if st.button("💾 สร้างและดาวน์โหลด บันทึกการตรวจค้น/ตรวจยึด (Word)", type="primary", use_container_width=True):
    try:
        doc = DocxTemplate("template_search_seizure.docx")
        
        context = {
            "record_location": record_location,
            "record_date_th": format_thai_date(record_date),
            "record_time": record_time,
            "search_location": search_location,
            "search_date_th": format_thai_date(search_date),
            "search_time": search_time,
            "search_end_time": search_end_time,
            "warrant_text": warrant_text,
            "leaders": leaders,
            "leaders_intro_text": leaders_intro_text,
            "units": units_data,
            "search_circumstances": search_circumstances,
            "seized_count": seized_count,
            "investigator_name": investigator_name,
            "img_signer": img_signer,
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
        st.error(f"เกิดข้อผิดพลาด: {e}")
