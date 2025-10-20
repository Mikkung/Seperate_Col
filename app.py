import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="แยกชีตตามอาจารย์ผู้สอน", page_icon="📘")

st.title("📘 แยกชีตในไฟล์ Excel ตามอาจารย์ผู้สอน (พร้อมสรุปท้ายตาราง)")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # หาชื่อคอลัมน์ "อาจารย์ผู้สอน"
    found_teacher_col = None
    for col in df.columns:
        if "อาจารย์" in str(col):
            found_teacher_col = col
            break

    if not found_teacher_col:
        st.error("❌ ไม่พบคอลัมน์ชื่อ 'อาจารย์ผู้สอน'")
    else:
        st.success(f"✅ พบคอลัมน์อาจารย์ผู้สอน: '{found_teacher_col}'")

        if st.button("🚀 สร้างไฟล์แยกชีตพร้อมสรุปท้ายตาราง"):
            output = BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for teacher, group in df.groupby(found_teacher_col):
                    group = group.copy()

                    # รีเซ็ท 'ลำดับ' ใหม่
                    first_col = df.columns[0]
                    group[first_col] = range(1, len(group) + 1)

                    # หาคอลัมน์จำนวนนิสิตและจำนวนเงิน (ค้นโดยคำ)
                    student_col = next((c for c in df.columns if "จำนวนนิสิต" in str(c)), None)
                    money_col = next((c for c in df.columns if "คิดเป็นเงิน" in str(c)), None)

                    # ✅ แปลงข้อมูลให้เป็นตัวเลขก่อน sum (ป้องกัน error)
                    if student_col:
                        group[student_col] = pd.to_numeric(group[student_col], errors='coerce').fillna(0)
                    if money_col:
                        # เอาเครื่องหมาย comma ออกก่อนแปลง เช่น "1,000" → 1000
                        group[money_col] = group[money_col].astype(str).str.replace(',', '', regex=False)
                        group[money_col] = pd.to_numeric(group[money_col], errors='coerce').fillna(0)

                    # ✅ คำนวณผลรวม
                    total_students = group[student_col].sum() if student_col else ""
                    total_money = group[money_col].sum() if money_col else ""

                    # ✅ เพิ่มแถวสรุปท้ายตาราง
                    summary = pd.DataFrame({
                        first_col: ["รวมเป็นเงิน"],
                        student_col: [total_students],
                        money_col: [total_money]
                    })

                    for col in group.columns:
                        if col not in summary.columns:
                            summary[col] = ""

                    summary = summary[group.columns]
                    final_df = pd.concat([group, summary], ignore_index=True)

                    safe_name = str(teacher).strip()[:31].replace('/', '-')
                    final_df.to_excel(writer, sheet_name=safe_name, index=False)

            output.seek(0)
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ที่แยกแล้ว (พร้อมสรุป)",
                data=output,
                file_name="แยกตามอาจารย์_พร้อมสรุป.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
