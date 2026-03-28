import streamlit as st
import pandas as pd
import requests
import json
import win32com.client as win32
import pythoncom 

# --- 1. ตั้งค่าหน้าจอ & Theme ---
st.set_page_config(page_title="CMS Maintenance System", layout="wide", page_icon="🚗")

# Custom CSS เพื่อความสวยงาม
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("🚗 ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)")
st.caption("จัดการการแจ้งเตือนผ่าน Outlook และ LINE อย่างมีประสิทธิภาพ")

# --- 2. Sidebar สำหรับคำแนะนำ (ปรับปรุงใหม่ให้อ่านง่ายขึ้น) ---
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

# --- 3. ส่วนการอัปโหลดไฟล์ ---
with st.container():
    col1, col2, col3 = st.columns(3)
    with col1:
        file_mileage = st.file_uploader("📂 1. รายงานการใช้รถ", type=['xlsx'])
    with col2:
        file_service = st.file_uploader("📂 2. ข้อมูลเข้าศูนย์", type=['xlsx'])
    with col3:
        file_config = st.file_uploader("📂 3. เงื่อนไข & API", type=['xlsx'])

# --- 4. ฟังก์ชันการทำงาน ---

def open_outlook_draft(to_addr, cc_addr, html_content, name):
    try:
        pythoncom.CoInitialize() 
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = str(to_addr) if pd.notna(to_addr) and str(to_addr) != "None" else ""
        mail.CC = str(cc_addr) if pd.notna(cc_addr) and str(cc_addr) != "None" else ""
        mail.Subject = f"📢 [แจ้งเตือน] กำหนดนำรถเข้าศูนย์บริการ: คุณ {name}"
        mail.HTMLBody = html_content
        mail.Display() 
        return True
    except Exception as e:
        return str(e)

def generate_individual_html(row):
    """สร้าง Email HTML ที่ดูหรูหราและเป็นทางการ"""
    is_over = row['ระยะห่าง'] < 0
    accent_color = "#d9534f" if is_over else "#f0ad4e"
    status_text = "เกินกำหนด" if is_over else "ใกล้ถึงกำหนด"
    
    return f"""
    <html>
    <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6;">
        <div style="max-width: 600px; border: 1px solid #eee; padding: 20px; border-radius: 10px;">
            <h2 style="color: #004d99; border-bottom: 2px solid #004d99; padding-bottom: 10px;">🚗 แจ้งเตือนการบำรุงรักษา</h2>
            <p>เรียน คุณ <strong>{row['ชื่อ-นามสกุล']}</strong>,</p>
            <p>ระบบ CMS ตรวจสอบพบว่ารถยนต์ของท่านถึงกำหนดตรวจสอบสภาพ ดังนี้:</p>
            
            <div style="background-color: {accent_color}; color: white; padding: 15px; border-radius: 5px; text-align: center; margin: 20px 0;">
                <span style="font-size: 18px; font-weight: bold;">สถานะ: {status_text}</span><br>
                <span style="font-size: 24px;">ระยะคงเหลือ {row['ระยะห่าง']:,} กม.</span>
            </div>

            <table style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>ทะเบียนรถ:</strong></td>
                    <td style="padding: 8px; border-bottom: 1px solid #eee;">{row['ป้ายทะเบียนรถ']}</td>
                </tr>
                <tr>
                    <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>เลขไมล์ปัจจุบัน:</strong></td>
                    <td style="padding: 8px; border-bottom: 1px solid #eee;">{row['เลขไมล์สิ้นสุด']:,} กม.</td>
                </tr>
                <tr>
                    <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>กำหนดเข้าศูนย์ที่:</strong></td>
                    <td style="padding: 8px; border-bottom: 1px solid #eee;">{row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']:,} กม.</td>
                </tr>
            </table>
            
            <p style="margin-top: 25px;">💡 <em>กรุณานัดหมายศูนย์บริการล่วงหน้าเพื่อความสะดวกของท่าน</em></p>
            <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
            <p style="font-size: 11px; color: #888; text-align: center;">นี่คือระบบแจ้งเตือนอัตโนมัติจาก CMS Maintenance System</p>
        </div>
    </body>
    </html>
    """

# --- 5. เริ่มต้นประมวลผล ---
if file_mileage and file_service and file_config:
    try:
        # --- อ่านข้อมูล ---
        with st.status("🚀 กำลังวิเคราะห์ข้อมูล...", expanded=False) as status:
            # LINE API
            line_token, line_user_id = "", ""
            try:
                df_line = pd.read_excel(file_config, sheet_name='LineAPI')
                if not df_line.empty:
                    line_token = str(df_line.iloc[0, 0]).strip()
                    line_user_id = str(df_line.iloc[0, 1]).strip()
            except: pass

            # Email Config
            df_e = pd.read_excel(file_config, sheet_name='เงื่อนไข')
            df_e.columns = df_e.columns.str.strip()
            df_e['Name'] = df_e['Name'].astype(str).str.strip()

            # Mileage
            df_m = pd.read_excel(file_mileage, header=2)
            df_m.columns = df_m.columns.str.strip()
            df_m['ชื่อ-นามสกุล'] = df_m['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
            last_m = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('ชื่อ-นามสกุล', as_index=False).last()

            # Service
            df_s = pd.read_excel(file_service, header=2)
            df_s.columns = df_s.columns.str.strip()
            df_s['ชื่อพนักงานขับรถปัจจุบัน'] = df_s['ชื่อพนักงานขับรถปัจจุบัน'].astype(str).str.strip()
            df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] = pd.to_numeric(df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'].astype(str).str.replace(',', ''), errors='coerce')

            # Process
            combined = pd.merge(last_m[['ชื่อ-นามสกุล', 'เลขไมล์สิ้นสุด']], 
                                df_s[['ชื่อพนักงานขับรถปัจจุบัน', 'ป้ายทะเบียนรถ', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป']], 
                                how='inner', left_on='ชื่อ-นามสกุล', right_on='ชื่อพนักงานขับรถปัจจุบัน')
            combined['ระยะห่าง'] = combined['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] - combined['เลขไมล์สิ้นสุด']
            alerts = pd.merge(combined, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')
            alerts = alerts[alerts['ระยะห่าง'] <= 500].sort_values('ระยะห่าง').copy()
            status.update(label="✅ ตรวจสอบข้อมูลเสร็จสิ้น", state="complete")

        # --- 6. แสดงผลลัพธ์ ---
        if not alerts.empty:
            # Metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("รถที่ต้องดูแล", f"{len(alerts)} คัน")
            m2.metric("เกินกำหนด (🔴)", f"{len(alerts[alerts['ระยะห่าง'] < 0])} คัน")
            m3.metric("ใกล้ถึงกำหนด (🟡)", f"{len(alerts[alerts['ระยะห่าง'] >= 0])} คัน")

            st.write("---")
            
            # ตารางแสดงผล
            def color_status(val):
                color = '#ff4b4b' if val < 0 else '#ffa500'
                return f'color: {color}; font-weight: bold'

            styled_df = alerts[['ชื่อ-นามสกุล', 'ป้ายทะเบียนรถ', 'ระยะห่าง', 'to']].style.applymap(color_status, subset=['ระยะห่าง'])
            st.dataframe(styled_df, use_container_width=True)
            
            # ปุ่มดำเนินการ
            c1, c2 = st.columns(2)
            with c1:
                if st.button("📧 เปิด Preview Outlook แยกรายคน"):
                    count = 0
                    for _, row in alerts.iterrows():
                        html = generate_individual_html(row)
                        if open_outlook_draft(row['to'], row['CC'], html, row['ชื่อ-นามสกุล']):
                            count += 1
                    st.toast(f"เตรียม Email เสร็จแล้ว {count} ฉบับ", icon="📩")

            with c2:
                if st.button("💬 ส่ง LINE สรุปให้แอดมิน"):
                    if line_token and line_user_id:
                        bubbles = []
                        for _, row in alerts.iterrows():
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
                        payload = {"to": line_user_id, "messages": [{"type": "flex", "altText": "แจ้งเตือน CMS", "contents": {"type": "carousel", "contents": bubbles[:12]}}]}
                        if requests.post("https://api.line.me/v2/bot/message/push", headers=headers, data=json.dumps(payload)).status_code == 200:
                            st.toast("ส่ง LINE เรียบร้อยแล้ว", icon="✅")
                    else:
                        st.error("กรุณาตรวจสอบค่า API ในไฟล์ Excel")
        else:
            st.success("✨ ยอดเยี่ยม! รถทุกคันอยู่ในสภาพปกติ ไม่ต้องแจ้งเตือนในขณะนี้")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดทางเทคนิค: {e}")