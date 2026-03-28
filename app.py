import streamlit as st
import pandas as pd
import requests
import json
import urllib.parse
from datetime import datetime

# --- 1. ตั้งค่าหน้าจอ & Theme ---
st.set_page_config(page_title="CMS Maintenance System", layout="wide", page_icon="🚗")

# Custom CSS สำหรับความสวยงามและ Footer
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .card-container {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #eee;
        margin-bottom: 10px;
    }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f1f1;
        color: #555;
        text-align: center;
        padding: 5px;
        font-size: 12px;
        border-top: 1px solid #ddd;
        z-index: 999;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🚗 ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)")
st.caption("จัดการการแจ้งเตือนผ่าน Mail Link และ LINE อย่างมีประสิทธิภาพ")

# --- 2. Sidebar สำหรับคำแนะนำ (ดึงของเดิมกลับมา) ---
with st.sidebar:
    st.header("📖 คู่มือการใช้งาน")
    
    st.subheader("🛠 ขั้นตอนที่ 1: เตรียมไฟล์")
    st.info("""
    **ไฟล์ที่ 1 และ 2 (ข้อมูลรถ):**
    * ต้องมีหัวตารางอยู่ที่ **แถวที่ 3**
    
    **ไฟล์ที่ 3 (การตั้งค่า):**
    * **Sheet 'เงื่อนไข':** ใส่รายชื่อพนักงาน และ Email (To, CC)
    * **Sheet 'LineAPI':** ใส่รหัส LINE Token และ User ID
    """)

    st.subheader("⚠️ ข้อควรระวัง")
    st.warning("""
    **ชื่อพนักงาน:** ในทุกไฟล์ต้องพิมพ์ให้ **ตรงกันทุกตัวอักษร** (รวมถึงช่องว่าง) หากพิมพ์ผิดระบบจะหาข้อมูลไม่เจอ
    """)
    
    st.divider()
    st.write("✅ **สถานะระบบ:** พร้อมประมวลผล")
    st.write("💻 **Developer:** ITsupportR4")

# --- 3. ส่วนการอัปโหลดไฟล์ ---
with st.container():
    col1, col2, col3 = st.columns(3)
    with col1:
        file_mileage = st.file_uploader("📂 1. รายงานการใช้รถ", type=['xlsx'])
    with col2:
        file_service = st.file_uploader("📂 2. ข้อมูลเข้าศูนย์", type=['xlsx'])
    with col3:
        file_config = st.file_uploader("📂 3. เงื่อนไข & API", type=['xlsx'])

# --- 4. ฟังก์ชันสร้าง Mail Link ---
def create_mailto_link(row):
    to_addr = str(row['to']) if pd.notna(row['to']) else ""
    cc_addr = str(row['CC']) if pd.notna(row['CC']) else ""
    subject = f"📢 [แจ้งเตือน] กำหนดนำรถเข้าศูนย์บริการ: คุณ {row['ชื่อ-นามสกุล']}"
    body = f"เรียน คุณ {row['ชื่อ-นามสกุล']},\n\nระบบตรวจพบว่ารถทะเบียน {row['ป้ายทะเบียนรถ']} ถึงกำหนดเข้าศูนย์\nเลขไมล์ปัจจุบัน: {row['เลขไมล์สิ้นสุด']:,} กม.\nกำหนดเข้าศูนย์ที่: {row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']:,} กม.\nระยะคงเหลือ: {row['ระยะห่าง']:,} กม.\n\nกรุณานัดหมายศูนย์บริการล่วงหน้าหากมีข้อสงสัยหรือต้องการข้อมูลเพิ่มเติม กรุณาติดต่อฝ่ายทรัพยากรบุคคล ประจำพื้นที่ \n\nส่งจากระบบ CMS"
    return f"mailto:{to_addr}?cc={cc_addr}&subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"

# --- 5. เริ่มต้นประมวลผล ---
if file_mileage and file_service and file_config:
    try:
        with st.status("🚀 กำลังวิเคราะห์ข้อมูล...", expanded=False) as status:
            # LINE API
            df_line = pd.read_excel(file_config, sheet_name='LineAPI')
            line_token = str(df_line.iloc[0, 0]).strip() if not df_line.empty else ""
            line_user_id = str(df_line.iloc[0, 1]).strip() if not df_line.empty else ""

            # Email Config
            df_e = pd.read_excel(file_config, sheet_name='เงื่อนไข')
            df_e.columns = df_e.columns.str.strip()
            df_e['Name'] = df_e['Name'].astype(str).str.strip()

            # Mileage (Header row 3)
            df_m = pd.read_excel(file_mileage, header=2)
            df_m.columns = df_m.columns.str.strip()
            df_m['ชื่อ-นามสกุล'] = df_m['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
            last_m = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('ชื่อ-นามสกุล', as_index=False).last()

            # Service (Header row 3)
            df_s = pd.read_excel(file_service, header=2)
            df_s.columns = df_s.columns.str.strip()
            df_s['ชื่อพนักงานขับรถปัจจุบัน'] = df_s['ชื่อพนักงานขับรถปัจจุบัน'].astype(str).str.strip()
            df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] = pd.to_numeric(df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'].astype(str).str.replace(',', ''), errors='coerce')

            # Process Merge
            combined = pd.merge(last_m[['ชื่อ-นามสกุล', 'เลขไมล์สิ้นสุด']], 
                                df_s[['ชื่อพนักงานขับรถปัจจุบัน', 'ป้ายทะเบียนรถ', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป']], 
                                how='inner', left_on='ชื่อ-นามสกุล', right_on='ชื่อพนักงานขับรถปัจจุบัน')
            combined['ระยะห่าง'] = combined['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] - combined['เลขไมล์สิ้นสุด']
            alerts = pd.merge(combined, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')
            
            # กรองพนักงานที่ต้องเข้าศูนย์ (ระยะห่าง <= 500)
            alerts = alerts[alerts['ระยะห่าง'] <= 500].sort_values('ระยะห่าง').copy()
            status.update(label="✅ ตรวจสอบข้อมูลเสร็จสิ้น", state="complete")

        # --- 6. แสดงผลลัพธ์ ---
        if not alerts.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("รถที่ต้องดูแล", f"{len(alerts)} คัน")
            m2.metric("เกินกำหนด (🔴)", f"{len(alerts[alerts['ระยะห่าง'] < 0])} คัน")
            m3.metric("ใกล้ถึงกำหนด (🟡)", f"{len(alerts[alerts['ระยะห่าง'] >= 0])} คัน")

            st.write("---")
            
            # --- ปุ่ม Export พนักงานที่ต้องเข้าศูนย์ ---
            st.subheader("💾 ส่งออกรายงาน (Export)")
            export_df = alerts[['ชื่อ-นามสกุล', 'ป้ายทะเบียนรถ', 'ระยะห่าง', 'เลขไมล์สิ้นสุด', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป', 'to']].copy()
            export_df['ผู้จัดทำ'] = "บุคคลและธุรการป้อมนำR4"
            csv = export_df.to_csv(index=False).encode('utf-8-sig')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์พนักงานที่ต้องเข้าศูนย์ (CSV)",
                data=csv,
                file_name=f'CMS_Alert_Report_{datetime.now().strftime("%Y%m%d")}.csv',
                mime='text/csv',
                use_container_width=True
            )

            st.write("---")
            st.subheader("📩 รายการแจ้งเตือนรายคน")
            for _, row in alerts.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="card-container">
                        <strong>👤 {row['ชื่อ-นามสกุล']}</strong> | 🚗 ทะเบียน: {row['ป้ายทะเบียนรถ']}<br>
                        เลขไมล์ปัจจุบัน: {row['เลขไมล์สิ้นสุด']:,} กม. | ระยะคงเหลือ: <span style="color:{'red' if row['ระยะห่าง'] < 0 else 'orange'}">{row['ระยะห่าง']:,} กม.</span>
                    </div>
                    """, unsafe_allow_html=True)
                    st.link_button(f"📧 เปิด Mail Link หา {row['ชื่อ-นามสกุล'].split()[0]}", create_mailto_link(row), use_container_width=True)

            # --- ปุ่มส่ง LINE ---
            st.write("---")
            if st.button("💬 ส่ง LINE สรุปให้แอดมิน", type="primary", use_container_width=True):
                if line_token and line_user_id:
                    bubbles = []
                    for _, row in alerts.iterrows()[:12]:
                        status_ico = "🔴" if row['ระยะห่าง'] < 0 else "🟡"
                        bubbles.append({
                            "type": "bubble", "size": "micro",
                            "body": {"type": "box", "layout": "vertical", "contents": [
                                {"type": "text", "text": f"{status_ico} {row['ป้ายทะเบียนรถ']}", "weight": "bold", "size": "sm"},
                                {"type": "text", "text": f"{row['ระยะห่าง']:,} กม.", "weight": "bold", "color": "#ff4b4b" if row['ระยะห่าง'] < 0 else "#ffa500"}
                            ]}
                        })
                    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {line_token}"}
                    payload = {"to": line_user_id, "messages": [{"type": "flex", "altText": "CMS Alert", "contents": {"type": "carousel", "contents": bubbles}}]}
                    if requests.post("https://api.line.me/v2/bot/message/push", headers=headers, data=json.dumps(payload)).status_code == 200:
                        st.toast("ส่ง LINE เรียบร้อยแล้ว", icon="✅")
        else:
            st.success("✨ ยอดเยี่ยม! รถทุกคันอยู่ในสภาพปกติ ไม่ต้องแจ้งเตือนในขณะนี้")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดทางเทคนิค: {e}")

# --- Footer เครดิต ---
st.markdown("""
    <div class="footer">
        Developed by <b>ITsupportR4</b> | CMS Maintenance System v1.0
    </div>
    """, unsafe_allow_html=True)
