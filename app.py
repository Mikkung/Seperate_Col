import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="แยกชีตตามอาจารย์ผู้สอน", page_icon="📘")

st.title("📘 แยกชีตในไฟล์ Excel ตามอาจารย์ผู้สอน (พร้อมสรุปและ merge cell)")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # ✅ ถ้า column แรกไม่มีชื่อ -> ตั้งชื่อให้ว่า "ลำดับ"
    if not str(df.columns[0]).strip() or "unnamed" in str(df.columns[0]).lower():
        df.columns = ["ลำดับ"] + list(df.columns[1:])
    else:
        df.rename(columns={df.columns[0]: "ลำดับ"}, inplace=True)

    # หาชื่อคอลัมน์ "อาจารย์ผู้สอน"
    teacher_col = next((c for c in df.columns if "อาจารย์" in str(c)), None)

    if not teacher_col:
        st.error("❌ ไม่พบคอลัมน์ชื่อ 'อาจารย์ผู้สอน'")
    else:
        st.success(f"✅ พบคอลัมน์อาจารย์ผู้สอน: '{teacher_col}'")

        if st.button("🚀 สร้างไฟล์แยกชีตพร้อมสรุป"):
            output = BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for teacher, group in df.groupby(teacher_col):
                    group = group.copy()

                    # 🧹 ตัดคอลัมน์สุดท้ายออก
                    group = group.iloc[:, :-1]

                    # ✅ ตั้งชื่อคอลัมน์แรกให้แน่ใจว่าเป็น "ลำดับ"
                    group.rename(columns={group.columns[0]: "ลำดับ"}, inplace=True)

                    # รีเซ็ตลำดับใหม่
                    group["ลำดับ"] = range(1, len(group) + 1)

                    # หาคอลัมน์จำนวนนิสิตและจำนวนเงิน
                    student_col = next((c for c in group.columns if "นิสิต" in str(c)), None)
                    money_col = next((c for c in group.columns if "เงิน" in str(c)), None)

                    # แปลงข้อมูลให้เป็นตัวเลข
                    if student_col:
                        group[student_col] = pd.to_numeric(group[student_col], errors='coerce').fillna(0)
                    if money_col:
                        group[money_col] = group[money_col].astype(str).str.replace(',', '', regex=False)
                        group[money_col] = pd.to_numeric(group[money_col], errors='coerce').fillna(0)

                    # รวมค่า
                    total_students = group[student_col].sum() if student_col else ""
                    total_money = group[money_col].sum() if money_col else ""

                    # เพิ่มแถวสรุป
                    summary = pd.DataFrame({
