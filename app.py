import streamlit as st
import pandas as pd
import requests
import json
import urllib.parse  # สำหรับจัดการภาษาไทยใน Mail Link

# --- 1. ตั้งค่าหน้าจอ & Theme ---
st.set_page_config(page_title="CMS Maintenance System", layout="wide", page_icon="🚗")

# Custom CSS
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
    </style>
    """, unsafe_allow_html=True)

st.title("🚗 ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)")
st.caption("จัดการการแจ้งเตือนผ่าน Mail Link และ LINE อย่างมีประสิทธิภาพ (Cloud Version)")

# --- 2. Sidebar ---
with st.sidebar:
    st.header("📖 คู่มือการใช้งาน (Cloud)")
    st.info("""
    **ขั้นตอน:**
    1. อัปโหลดไฟล์ Excel ทั้ง 3 ชุด
    2. ตรวจสอบรายชื่อรถที่ต้องดูแลในตาราง
    3. กดปุ่ม '📧 ส่งเมล' เพื่อเปิดโปรแกรมเมลในเครื่องของคุณ
    4. กดปุ่ม '💬 ส่ง LINE' เพื่อแจ้งสรุปเข้ากลุ่มแอดมิน
    """)
    st.divider()
    st.write("✅ **สถานะ:** พร้อมใช้งานบน Cloud")

# --- 3. ส่วนการอัปโหลดไฟล์ ---
col1, col2, col3 = st.columns(3)
with col1:
    file_mileage = st.file_uploader("📂 1. รายงานการใช้รถ", type=['xlsx'])
with col2:
    file_service = st.file_uploader("📂 2. ข้อมูลเข้าศูนย์", type=['xlsx'])
with col3:
    file_config = st.file_uploader("📂 3. เงื่อนไข & API", type=['xlsx'])

# --- 4. ฟังก์ชันการทำงาน ---

def create_mailto_link(row):
    """สร้าง Link สำหรับเปิดโปรแกรม Email (รองรับภาษาไทย)"""
    to_addr = str(row['to']) if pd.notna(row['to']) else ""
    cc_addr = str(row['CC']) if pd.notna(row['CC']) else ""
    
    subject = f"📢 [แจ้งเตือน] กำหนดนำรถเข้าศูนย์บริการ: คุณ {row['ชื่อ-นามสกุล']}"
    
    # เนื้อหาเมล (แบบ Text ล้วน เพราะ Mailto ไม่รองรับ HTML ซับซ้อน)
    body = f"""เรียน คุณ {row['ชื่อ-นามสกุล']},

ระบบ CMS ตรวจสอบพบว่ารถยนต์ของท่านถึงกำหนดตรวจสอบสภาพ ดังนี้:

🚗 ทะเบียนรถ: {row['ป้ายทะเบียนรถ']}
📍 เลขไมล์ปัจจุบัน: {row['เลขไมล์สิ้นสุด']:,} กม.
🛠️ กำหนดเข้าศูนย์ที่: {row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']:,} กม.
⚠️ ระยะคงเหลือ: {row['ระยะห่าง']:,} กม.

กรุณานัดหมายศูนย์บริการล่วงหน้าเพื่อความสะดวกของท่าน

----------------------------------
ส่งจากระบบอัตโนมัติ CMS Maintenance System
    """
    
    # เข้ารหัสข้อความเพื่อให้ URL อ่านภาษาไทยออก
    subject_enc = urllib.parse.quote(subject)
    body_enc = urllib.parse.quote(body)
    
    return f"mailto:{to_addr}?cc={cc_addr}&subject={subject_enc}&body={body_enc}"

# --- 5. เริ่มต้นประมวลผล ---
if file_mileage and file_service and file_config:
    try:
        with st.status("🚀 กำลังวิเคราะห์ข้อมูล...", expanded=False) as status:
            # LINE API Config
            df_line = pd.read_excel(file_config, sheet_name='LineAPI')
            line_token = str(df_line.iloc[0, 0]).strip() if not df_line.empty else ""
            line_user_id = str(df_line.iloc[0, 1]).strip() if not df_line.empty else ""

            # Email Config
            df_e = pd.read_excel(file_config, sheet_name='เงื่อนไข')
            df_e.columns = df_e.columns.str.strip()
            df_e['Name'] = df_e['Name'].astype(str).str.strip()

            # Mileage Data (Header แถวที่ 3)
            df_m = pd.read_excel(file_mileage, header=2)
            df_m.columns = df_m.columns.str.strip()
            df_m['ชื่อ-นามสกุล'] = df_m['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
            last_m = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('ชื่อ-นามสกุล', as_index=False).last()

            # Service Data (Header แถวที่ 3)
            df_s = pd.read_excel(file_service, header=2)
            df_s.columns = df_s.columns.str.strip()
            df_s['ชื่อพนักงานขับรถปัจจุบัน'] = df_s['ชื่อพนักงานขับรถปัจจุบัน'].astype(str).str.strip()
            df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] = pd.to_numeric(df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'].astype(str).str.replace(',', ''), errors='coerce')

            # Merge Data
            combined = pd.merge(last_m[['ชื่อ-นามสกุล', 'เลขไมล์สิ้นสุด']], 
                                df_s[['ชื่อพนักงานขับรถปัจจุบัน', 'ป้ายทะเบียนรถ', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป']], 
                                how='inner', left_on='ชื่อ-นามสกุล', right_on='ชื่อพนักงานขับรถปัจจุบัน')
            
            combined['ระยะห่าง'] = combined['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] - combined['เลขไมล์สิ้นสุด']
            alerts = pd.merge(combined, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')
            
            # กรองเฉพาะที่เหลือไม่เกิน 500 กม.
            alerts = alerts[alerts['ระยะห่าง'] <= 500].sort_values('ระยะห่าง').copy()
            status.update(label="✅ ตรวจสอบข้อมูลเสร็จสิ้น", state="complete")

        # --- 6. แสดงผลลัพธ์ ---
        if not alerts.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("รถที่ต้องดูแล", f"{len(alerts)} คัน")
            m2.metric("เกินกำหนด (🔴)", f"{len(alerts[alerts['ระยะห่าง'] < 0])} คัน")
            m3.metric("ใกล้ถึงกำหนด (🟡)", f"{len(alerts[alerts['ระยะห่าง'] >= 0])} คัน")

            st.write("---")
            st.subheader("📋 รายละเอียดการแจ้งเตือน")

            for index, row in alerts.iterrows():
                with st.container():
                    # สร้าง UI แบบ Card
                    st.markdown(f"""
                    <div class="card-container">
                        <strong>👤 {row['ชื่อ-นามสกุล']}</strong> | 🚗 ทะเบียน: {row['ป้ายทะเบียนรถ']}<br>
                        ระยะคงเหลือ: <span style="color:{'red' if row['ระยะห่าง'] < 0 else 'orange'}">{row['ระยะห่าง']:,} กม.</span>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    c1, c2 = st.columns([1, 1])
                    with c1:
                        # ปุ่มเปิดโปรแกรมเมล (Mailto)
                        mail_url = create_mailto_link(row)
                        st.link_button(f"📧 ส่งอีเมลหา {row['ชื่อ-นามสกุล'].split()[0]}", mail_url, use_container_width=True)
                    
                    st.write("") # ระยะห่างเล็กน้อย

            st.write("---")
            # ปุ่มส่ง LINE สรุปทั้งหมดให้แอดมิน
            if st.button("💬 ส่ง LINE สรุปภาพรวมให้แอดมิน", type="primary", use_container_width=True):
                if line_token and line_user_id:
                    bubbles = []
                    for _, row in alerts.iterrows()[:12]: # LINE Flex Carousel รับได้สูงสุด 12 ใบ
                        status_ico = "🔴" if row['ระยะห่าง'] < 0 else "🟡"
                        bubbles.append({
                            "type": "bubble", "size": "micro",
                            "body": {"type": "box", "layout": "vertical", "contents": [
                                {"type": "text", "text": f"{status_ico} {row['ป้ายทะเบียนรถ']}", "weight": "bold", "size": "sm"},
                                {"type": "text", "text": f"{row['ชื่อ-นามสกุล']}", "size": "xs", "color": "#666666"},
                                {"type": "text", "text": f"{row['ระยะห่าง']:,} กม.", "weight": "bold", "color": "#ff4b4b" if row['ระยะห่าง'] < 0 else "#ffa500"}
                            ]}
                        })
                    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {line_token}"}
                    payload = {"to": line_user_id, "messages": [{"type": "flex", "altText": "CMS Alert", "contents": {"type": "carousel", "contents": bubbles}}]}
                    if requests.post("https://api.line.me/v2/bot/message/push", headers=headers, data=json.dumps(payload)).status_code == 200:
                        st.toast("ส่ง LINE เรียบร้อยแล้ว", icon="✅")
                else:
                    st.error("ไม่พบข้อมูล LINE Token ในไฟล์ตั้งค่า")
        else:
            st.success("✨ ยอดเยี่ยม! รถทุกคันอยู่ในสภาพปกติ")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
