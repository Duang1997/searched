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

# --- ฟังก์ชันจัดการวันที่ ---
THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def format_thai_date(date_obj):
    if not date_obj: return ""
    return f"วันที่ {date_obj.day} {THAI_MONTHS[date_obj.month]} {date_obj.year + 543}"

# --- ฟอร์มบันทึกการตรวจค้น/ตรวจยึด ---
st.header("ส่วนที่ 1: ข้อมูลวันเวลาและสถานที่")
col1, col2 = st.columns(2)
with col1:
    record_location = st.text_input("สถานที่บันทึก", placeholder="เช่น ร้านโอเครถเช่า...")
    record_date = st.date_input("วันที่บันทึก")
    record_time = st.text_input("เวลาที่บันทึก", placeholder="เช่น 08.30")
with col2:
    search_location = st.text_input("สถานที่ตรวจค้น", placeholder="เช่น ร้านโอเครถเช่า...")
    search_date = st.date_input("วันที่ตรวจค้น/ตรวจยึด")
    search_time = st.text_input("เวลาที่ตรวจค้น/ตรวจยึด", placeholder="เช่น 06.00")
    search_end_time = st.text_input("เวลาที่เสร็จสิ้นการตรวจค้น", placeholder="เช่น 08.00")

st.divider()

st.header("ส่วนที่ 2: อำนาจในการตรวจค้นและผู้นำตรวจค้น")
col3, col4 = st.columns(2)
with col3:
    st.subheader("รายละเอียดหมายค้น")
    warrant_court = st.text_input("ศาลที่ออกหมายค้น", placeholder="เช่น ศาลจังหวัดอุบลราชธานี")
    warrant_no = st.text_input("หมายค้นที่", placeholder="เช่น 47/2568")
    warrant_date = st.date_input("ลงวันที่ (หมายค้น)")
with col4:
    st.subheader("เจ้าบ้าน/ผู้นำตรวจค้น")
    leader_name_1 = st.text_input("ชื่อ-นามสกุล ผู้นำตรวจค้น (1)", placeholder="ระบุชื่อ...")
    leader_status_1 = st.text_input("สถานะ (1)", placeholder="เช่น ผู้ดูแลสถานที่, เจ้าบ้าน")
    leader_name_2 = st.text_input("ชื่อ-นามสกุล ผู้นำตรวจค้น (2)", placeholder="ระบุชื่อ (ถ้ามี)...")

st.divider()

st.header("ส่วนที่ 3: ข้อมูลเจ้าพนักงานตำรวจ")
st.text_area("ภายใต้การอำนวยการสั่งการของ", placeholder="พล.ต.ท....")
st.text_area("เจ้าพนักงานตำรวจชุดตรวจค้น ประกอบด้วย", placeholder="ระบุรายชื่อเจ้าพนักงานผู้ปฏิบัติการ...")

st.divider()

st.header("ส่วนที่ 4: พฤติการณ์และผลการตรวจค้น")
search_circumstances = st.text_area("พฤติการณ์ในการตรวจค้น/ตรวจยึด", height=150, placeholder="ตามวันเวลาที่แจ้ง เจ้าพนักงานตำรวจชุดตรวจค้น ได้นำหมายค้นศาล...")
investigator_name = st.text_input("พนักงานสอบสวนผู้รับผิดชอบ", placeholder="นำทรัพย์ทั้งหมดส่ง พงส. ...")

st.subheader("รายการสิ่งของตรวจยึด")
seized_df = pd.DataFrame([{"ลำดับ": "1", "รายการสิ่งของ": "", "จำนวน": "", "หมายเหตุ": ""}])
edited_seized = st.data_editor(seized_df, num_rows="dynamic", use_container_width=True)

st.divider()

st.header("ส่วนที่ 5: ภาพประกอบการตรวจค้น")
st.caption("อัปโหลดภาพถ่ายการปฏิบัติงานเพื่อแนบท้ายบันทึก")
col_img1, col_img2 = st.columns(2)
with col_img1:
    img_1 = st.file_uploader("ภาพประกอบที่ 1", type=['png', 'jpg', 'jpeg'])
with col_img2:
    img_2 = st.file_uploader("ภาพประกอบที่ 2", type=['png', 'jpg', 'jpeg'])

st.divider()

if st.button("💾 สร้างและดาวน์โหลด บันทึกการตรวจค้น/ตรวจยึด", type="primary", use_container_width=True):
    # เตรียมข้อมูลสำหรับส่งเข้า DocxTemplate (ต้องสร้าง template_search_seizure.docx ให้ตรงกับตัวแปร)
    context = {
        "record_location": record_location,
        "record_date_th": format_thai_date(record_date),
        "record_time": record_time,
        "search_location": search_location,
        "search_date_th": format_thai_date(search_date),
        "search_time": search_time,
        "search_end_time": search_end_time,
        "warrant_court": warrant_court,
        "warrant_no": warrant_no,
        "warrant_date_th": format_thai_date(warrant_date),
        "leader_name_1": leader_name_1,
        "leader_name_2": leader_name_2,
        "search_circumstances": search_circumstances,
        "investigator_name": investigator_name,
        "seized_items": edited_seized.to_dict('records')
    }
    
    st.success("✅ ประมวลผลข้อมูลสำเร็จ (กรุณาเชื่อมต่อกับ DocxTemplate เพื่อส่งออกเอกสาร)")
    # Code สำหรับเรนเดอร์เอกสาร
    # doc = DocxTemplate("template_search_seizure.docx")
    # doc.render(context)
    # bio = BytesIO()
    # doc.save(bio)
    # st.download_button("⬇️ โหลดไฟล์ บันทึกตรวจค้น.docx", data=bio.getvalue(), file_name="บันทึกตรวจค้น.docx")