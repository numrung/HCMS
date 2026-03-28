import streamlit as st
import pandas as pd
import requests
import json
import urllib.parse
from datetime import datetime

# --- 1. ตั้งค่าหน้าจอ & Theme ---
st.set_page_config(page_title="CMS Maintenance System", layout="wide", page_icon="🚗")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f1f1;
        color: #555;
        text-align: center;
        padding: 10px;
        font-size: 14px;
        border-top: 1px solid #ddd;
    }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .card-container {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #eee;
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🚗 ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)")
st.caption("จัดการการแจ้งเตือนและส่งออกรายงานพนักงานที่ต้องเข้าศูนย์")

# --- 2. ส่วนการอัปโหลดไฟล์ ---
col1, col2, col3 = st.columns(3)
with col1:
    file_mileage = st.file_uploader("📂 1. รายงานการใช้รถ", type=['xlsx'])
with col2:
    file_service = st.file_uploader("📂 2. ข้อมูลเข้าศูนย์", type=['xlsx'])
with col3:
    file_config = st.file_uploader("📂 3. เงื่อนไข & API", type=['xlsx'])

# --- 3. ฟังก์ชันสร้าง Mail Link ---
def create_mailto_link(row):
    to_addr = str(row['to']) if pd.notna(row['to']) else ""
    cc_addr = str(row['CC']) if pd.notna(row['CC']) else ""
    subject = f"📢 [แจ้งเตือน] กำหนดนำรถเข้าศูนย์บริการ: คุณ {row['ชื่อ-นามสกุล']}"
    body = f"เรียน คุณ {row['ชื่อ-นามสกุล']},\n\nรถทะเบียน {row['ป้ายทะเบียนรถ']} ของท่านถึงกำหนดเข้าศูนย์\nระยะคงเหลือ: {row['ระยะห่าง']:,} กม.\n\nส่งจากระบบ CMS"
    return f"mailto:{to_addr}?cc={cc_addr}&subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"

# --- 4. ประมวลผลข้อมูล ---
if file_mileage and file_service and file_config:
    try:
        # อ่านข้อมูล (คงเดิมจาก Logic เก่าของคุณ)
        df_line = pd.read_excel(file_config, sheet_name='LineAPI')
        line_token = str(df_line.iloc[0, 0]).strip() if not df_line.empty else ""
        line_user_id = str(df_line.iloc[0, 1]).strip() if not df_line.empty else ""

        df_e = pd.read_excel(file_config, sheet_name='เงื่อนไข')
        df_e.columns = df_e.columns.str.strip()
        df_e['Name'] = df_e['Name'].astype(str).str.strip()

        df_m = pd.read_excel(file_mileage, header=2)
        df_m['ชื่อ-นามสกุล'] = df_m['ชื่อ-นามสกุล'].astype(str).str.strip()
        df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
        last_m = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('ชื่อ-นามสกุล', as_index=False).last()

        df_s = pd.read_excel(file_service, header=2)
        df_s['ชื่อพนักงานขับรถปัจจุบัน'] = df_s['ชื่อพนักงานขับรถปัจจุบัน'].astype(str).str.strip()
        df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] = pd.to_numeric(df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'].astype(str).str.replace(',', ''), errors='coerce')

        combined = pd.merge(last_m[['ชื่อ-นามสกุล', 'เลขไมล์สิ้นสุด']], 
                            df_s[['ชื่อพนักงานขับรถปัจจุบัน', 'ป้ายทะเบียนรถ', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป']], 
                            how='inner', left_on='ชื่อ-นามสกุล', right_on='ชื่อพนักงานขับรถปัจจุบัน')
        
        combined['ระยะห่าง'] = combined['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] - combined['เลขไมล์สิ้นสุด']
        alerts = pd.merge(combined, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')
        
        # กรองเฉพาะคนที่ต้องเข้าศูนย์ (ระยะ <= 500 กม.)
        alerts = alerts[alerts['ระยะห่าง'] <= 500].sort_values('ระยะห่าง').copy()

        if not alerts.empty:
            st.success(f"พบพนักงานที่ต้องเข้าศูนย์ทั้งหมด {len(alerts)} ท่าน")
            
            # --- ส่วนปุ่ม Export (เฉพาะคนที่ต้องเข้าศูนย์) ---
            st.subheader("💾 ส่งออกรายงานพนักงานที่ต้องเข้าศูนย์")
            
            # เตรียมไฟล์สำหรับ Export และเพิ่มเครดิตในไฟล์
            export_data = alerts[['ชื่อ-นามสกุล', 'ป้ายทะเบียนรถ', 'เลขไมล์สิ้นสุด', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป', 'ระยะห่าง', 'to']].copy()
            export_data['Report_By'] = "ITsupportR4"
            export_data['Export_Date'] = datetime.now().strftime("%Y-%m-%d %H:%M")
            
            csv = export_data.to_csv(index=False).encode('utf-8-sig')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์สรุป (CSV) สำหรับคนที่จะแจ้งเตือน",
                data=csv,
                file_name=f'CMS_Maintenance_Alert_{datetime.now().strftime("%Y%m%d")}.csv',
                mime='text/csv',
                use_container_width=True
            )
            
            st.divider()

            # แสดงรายการรายคน (Card View)
            for _, row in alerts.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="card-container">
                        <strong>👤 {row['ชื่อ-นามสกุล']}</strong> | 🚗 {row['ป้ายทะเบียนรถ']}<br>
                        ระยะคงเหลือ: <span style="color:{'red' if row['ระยะห่าง'] < 0 else 'orange'}">{row['ระยะห่าง']:,} กม.</span>
                    </div>
                    """, unsafe_allow_html=True)
                    st.link_button(f"📧 ส่งเมลหา {row['ชื่อ-นามสกุล'].split()[0]}", create_mailto_link(row))

        else:
            st.success("✨ ยอดเยี่ยม! ไม่มีรถที่ต้องเข้าศูนย์ในขณะนี้")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")

# --- 5. ส่วนท้ายหน้าจอ (Footer) ---
st.markdown("""
    <div class="footer">
        <p>💻 Developed by <b>ITsupportR4</b> | CMS Maintenance System v1.0</p>
    </div>
    """, unsafe_allow_html=True)
