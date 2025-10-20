import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="แยกชีตตามอาจารย์ผู้สอน", page_icon="📘")

st.title("📘 แยกชีตในไฟล์ Excel ตามอาจารย์ผู้สอน")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    if 'อาจารย์ผู้สอน' not in df.columns:
        st.error("❌ ไม่พบคอลัมน์ชื่อ 'อาจารย์ผู้สอน' ในไฟล์ Excel")
    else:
        st.success("✅ โหลดไฟล์เรียบร้อย! คลิกปุ่มด้านล่างเพื่อสร้างไฟล์ใหม่")

        if st.button("🚀 สร้างไฟล์แยกชีต"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for teacher, group in df.groupby('อาจารย์ผู้สอน'):
                    safe_name = str(teacher).strip()[:31].replace('/', '-')
                    group.to_excel(writer, sheet_name=safe_name, index=False)
            
            output.seek(0)
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ที่แยกแล้ว",
                data=output,
                file_name="แยกตามอาจารย์.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
