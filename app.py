import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
import datetime

# --- ตั้งค่าหน้าเพจ (ยกเลิก layout="wide" เพื่อให้เป็นแนวตรงตามความต้องการ) ---
st.set_page_config(page_title="ระบบบันทึกการตรวจค้น/ตรวจยึด (CIB)") 

st.markdown("""
<style>
    .cib-header { background-color: #00204a; padding: 15px; border-radius: 5px; color: #f9bc0f; text-align: center; font-family: sans-serif; margin-bottom: 20px; }
    .cib-header h1 { color: #f9bc0f; margin: 0; font-size: 28px; font-weight: bold; }
    .cib-header p { color: #ffffff; margin: 0; font-size: 16px; }
</style>
<div class="cib-header">
    <h1>ตำรวจสอบสวนกลาง (Central Investigation Bureau)</h1>
    <p>ระบบจัดทำบันทึกการตรวจค้น / ตรวจยึด</p>
</div>
""", unsafe_allow_html=True)

# --- ฟังก์ชันจัดการวันที่ ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def format_thai_date(date_obj):
    if not date_obj: return ""
    return f"วันที่ {date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

def format_ad_date(date_obj):
    if not date_obj: return ""
    return date_obj.strftime("%d/%m/%Y")

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
warrant_court = st.text_input("ศาลที่ออกหมายค้น")
warrant_no = st.text_input("หมายค้นที่")
warrant_date = st.date_input("ลงวันที่ (หมายค้น)", key="warrant_date")

st.subheader("ผู้นำตรวจค้น")
leaders = []
for i in range(st.session_state.leader_count):
    with st.container(border=True):
        lname = st.text_input(f"ชื่อ-นามสกุล ผู้นำตรวจค้นคนที่ {i+1}", key=f"l_name_{i}")
        lstatus = st.text_input(f"สถานะ ผู้นำตรวจค้นคนที่ {i+1}", placeholder="เช่น ผู้ดูแลสถานที่, เจ้าบ้าน", key=f"l_status_{i}")
        if lname.strip():
            leaders.append({"name": lname.strip(), "status": lstatus.strip()})
st.button("➕ เพิ่มผู้นำตรวจค้น", on_click=add_leader)

st.divider()

# ==========================================
# ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ
# ==========================================
st.header("ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ")
default_cmd = "พล.ต.ต.ณัฐศักดิ์ เชาวนาศัย ผบช.ก., พ.ต.ต.พัฒนศักดิ์ บุบผาสุวรรณ ผบก.ป., พ.ต.อ.สุเทพ โตอิ้ม รอง ผบก.ป., พ.ต.อ.สุริยศักดิ์ จิราวัสน์ ผกก.3 บก.ป., พ.ต.ท.พงษ์พิทักษ์ เหล็กชูชาติ, พ.ต.ท.รัฐมนตรี พันชูกลาง, พ.ต.ท.ณัฐดนัย สีแข่ไตร, พ.ต.ท.ศิษฏ์ พูลวงศ์, พ.ต.ท.พัฒษพงศ์ เสณีแสนเสนา รอง ผกก.3 บก.ป."
commanders = st.text_area("ภายใต้อำนวยการสั่งการของ", value=default_cmd, height=100)

units_data_text = []
officers_data = []
officer_displays = []

st.subheader("หน่วยและเจ้าหน้าที่ที่ร่วมตรวจค้น")
for i in range(st.session_state.unit_count):
    with st.container(border=True):
        st.markdown(f"**หน่วยตรวจค้นที่ {i+1}**")
        unit_name = st.text_input(f"ชื่อหน่วยงานที่ {i+1}", placeholder="เช่น กก.3 บก.ป.", key=f"u_name_{i}")
        
        df_key = f"officer_df_{i}"
        if df_key not in st.session_state:
            st.session_state[df_key] = pd.DataFrame([{"ยศ": "พ.ต.ต.", "ชื่อ-นามสกุล": "", "ตำแหน่ง": "สว.กก.๓ บก.ป."}])
            
        off_mode = st.radio(f"รูปแบบการเพิ่มรายชื่อ (หน่วยที่ {i+1})", ["กรอกผ่านตารางในเว็บ", "อัปโหลดไฟล์ Excel"], horizontal=True, key=f"mode_{i}")
        
        if off_mode == "อัปโหลดไฟล์ Excel":
            up_file = st.file_uploader(f"อัปโหลดไฟล์ Excel (.xlsx) หน่วยที่ {i+1}", type=["xlsx"], key=f"up_{i}")
            if up_file is not None:
                if f"uploaded_{i}" not in st.session_state or st.session_state[f"uploaded_{i}"] != up_file.name:
                    df = pd.read_excel(up_file, dtype=str)
                    df.columns = df.columns.str.strip()
                    if "ชื่อ-นามสกุล" in df.columns:
                        st.session_state[df_key] = df
                        st.session_state[f"uploaded_{i}"] = up_file.name
                    else:
                        st.error("⚠️ ไม่พบคอลัมน์ 'ชื่อ-นามสกุล' ในไฟล์")

        # แสดงตารางเพื่อให้แก้ไขได้เสมอ ไม่ว่าจะมาจากการกรอกหรือ Excel
        edited_officers = st.data_editor(st.session_state[df_key], num_rows="dynamic", use_container_width=True, key=f"edit_{i}")
        
        unit_officers = []
        if "ชื่อ-นามสกุล" in edited_officers.columns:
            valid_officers = edited_officers[edited_officers["ชื่อ-นามสกุล"].astype(str).str.strip() != ""]
            valid_officers = valid_officers[valid_officers["ชื่อ-นามสกุล"].astype(str).str.lower() != "nan"]
            
            for _, r in valid_officers.iterrows():
                rank = str(r.get('ยศ', '')).replace("nan", "").strip()
                name = str(r.get('ชื่อ-นามสกุล', '')).replace("nan", "").strip()
                pos = str(r.get('ตำแหน่ง', '')).replace("nan", "").strip()
                display = f"{rank}{name} {pos}".strip()
                
                if name:
                    officer_obj = {"rank": rank, "name_only": name, "display": display}
                    unit_officers.append(officer_obj)
                    officers_data.append(officer_obj)
                    officer_displays.append(display)
                    
        if unit_name and unit_officers:
            units_data_text.append(f"เจ้าพนักงานตำรวจ ({unit_name}) ประกอบด้วย " + ", ".join([o['display'] for o in unit_officers]))

st.button("➕ เพิ่มหน่วยตรวจค้นอื่น", on_click=add_unit)
police_team_text = " และ ".join(units_data_text) if units_data_text else ", ".join(officer_displays)

signature_rows = []
for j in range(0, len(officers_data), 2):
    o1 = officers_data[j]
    o2 = officers_data[j+1] if j+1 < len(officers_data) else {"rank": "", "name_only": ""}
    signature_rows.append({
        "officer1_rank": o1["rank"], "officer1_name": o1["name_only"], 
        "officer2_rank": o2["rank"], "officer2_name": o2["name_only"]
    })

st.divider()

# ==========================================
# ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น
# ==========================================
st.header("ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น")
search_circumstances = st.text_area("พฤติการณ์ในการตรวจค้น/ตรวจยึด", height=150)
seized_count = st.number_input("จำนวนรายการสิ่งของตรวจยึด (รายการ)", min_value=0, value=1)

handover_opt = st.radio("นำทรัพย์ทั้งหมดส่งมอบให้ใคร", ["พนักงานสอบสวนผู้รับผิดชอบ", "อื่นๆ (ระบุ)"])
if handover_opt == "อื่นๆ (ระบุ)":
    investigator_name = st.text_input("ระบุชื่อ/ตำแหน่งผู้รับมอบ")
else:
    investigator_name = "พนักงานสอบสวนผู้รับผิดชอบ"

st.divider()

# ==========================================
# ส่วนที่ 5: ภาพประกอบการตรวจค้น
# ==========================================
st.header("ส่วนที่ 5: ภาพประกอบการตรวจค้น")
img_1 = st.file_uploader("ภาพประกอบที่ 1", type=['png', 'jpg', 'jpeg'])
img_2 = st.file_uploader("ภาพประกอบที่ 2", type=['png', 'jpg', 'jpeg'])

st.subheader("เลือกหัวหน้าชุดตรวจค้น")
leader_names = [l["name"] for l in leaders]
sign_choices = leader_names + officer_displays + ["อื่นๆ (ระบุ)"]

img_sign_opt = st.selectbox("เลือกหัวหน้าชุดตรวจค้น", sign_choices)
if img_sign_opt == "อื่นๆ (ระบุ)":
    img_signer = st.text_input("เลือกหัวหน้าชุดตรวจค้น")
else:
    img_signer = img_sign_opt

st.divider()

# ==========================================
# การประมวลผลและสร้างเอกสาร
# ==========================================
if st.button("💾 สร้างและดาวน์โหลด บันทึกการตรวจค้น/ตรวจยึด", type="primary", use_container_width=True):
    try:
        doc = DocxTemplate("template_search_seizure.docx")
        
        context = {
            "record_location": record_location,
            "record_date_th": format_thai_date(record_date), 
            "record_date_ad": format_ad_date(record_date),
            "record_time": record_time,
            "search_location": search_location,
            "search_date_th": format_thai_date(search_date),
            "search_date_ad": format_ad_date(search_date),
            "search_time": search_time,
            "search_end_time": search_end_time,
            "warrant_court": warrant_court,
            "warrant_no": warrant_no,
            "warrant_date_th": format_thai_date(warrant_date),
            "warrant_date_ad": format_ad_date(warrant_date),
            "leaders": leaders,
            "commanders": commanders,
            "police_team_text": police_team_text,
            "signature_rows": signature_rows,
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
