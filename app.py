import os
import json
import datetime
import urllib.parse
import pandas as pd
import streamlit as st

# --- 1. ตั้งค่าหน้าจอ & Theme ---
st.set_page_config(page_title="CMS Maintenance System", layout="wide", page_icon="🚗")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .card-container {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #eee;
        margin-bottom: 5px;
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

st.title("🚗 ระบบแจ้งเตือนบำรุงรักษา & เปลี่ยนยาง Auto-Fetch (CMS)")
st.caption("ระบบตรวจสอบระยะเข้าศูนย์ และอายุยางรถยนต์ ผ่านการอัปโหลดไฟล์ และกดส่งอีเมลผ่าน HTML Link (mailto)")

# --- 2. คำนวณหาชื่อไฟล์ประจำเดือนปัจจุบันตามเกณฑ์ (สำหรับแนะนำผู้ใช้งาน) ---
now = datetime.datetime.now()
months_th = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
current_month_name = months_th[now.month]
current_year_th = now.year + 543

expected_mileage_file = f"รายงานการใช้รถ_{current_month_name}_{current_year_th}.xlsx"
expected_service_file = f"รายงานการเข้าศูนย์_{current_month_name}_{current_year_th}.xlsx"

# --- 3. Sidebar: ระบบอัปโหลดไฟล์ (st.file_uploader) และตรวจเช็คสถานะ ---
with st.sidebar:
    st.header("📁 อัปโหลดไฟล์ระบบ (Excel)")
    
    uploaded_mileage = st.file_uploader(f"1. ไฟล์รายงานการใช้รถ ({expected_mileage_file})", type=["xlsx"])
    uploaded_service = st.file_uploader(f"2. ไฟล์ข้อมูลเข้าศูนย์ ({expected_service_file})", type=["xlsx"])
    uploaded_tyre = st.file_uploader("3. ไฟล์รายงานการเปลี่ยนยาง.xlsx", type=["xlsx"])
    uploaded_config = st.file_uploader("4. ไฟล์เงื่อนไข & API (Email.xlsx)", type=["xlsx"])
    
    st.divider()
    st.header("📖 Status สถานะไฟล์")
    
    file_mileage_ready = uploaded_mileage is not None
    file_service_ready = uploaded_service is not None
    file_tyre_ready = uploaded_tyre is not None
    file_config_ready = uploaded_config is not None

    if file_mileage_ready: st.success("🟢 โหลดไฟล์รายงานการใช้รถสำเร็จ")
    else: st.error("🔴 ยังไม่ได้อัปโหลดไฟล์รายงานการใช้รถ")

    if file_service_ready: st.success("🟢 โหลดไฟล์ข้อมูลเข้าศูนย์สำเร็จ")
    else: st.error("🔴 ยังไม่ได้อัปโหลดไฟล์ข้อมูลเข้าศูนย์")

    if file_tyre_ready: st.success("🟢 โหลดไฟล์ประวัติเปลี่ยนยางสำเร็จ")
    else: st.error("🔴 ยังไม่ได้อัปโหลดไฟล์ประวัติเปลี่ยนยาง")

    if file_config_ready: st.success("🟢 โหลดไฟล์เงื่อนไขสำเร็จ")
    else: st.error("🔴 ยังไม่ได้อัปโหลดไฟล์เงื่อนไข (Email.xlsx)")
        
    st.divider()
    st.write("💻 **Developer:** ITsupportR4")

# --- 4. ฟังก์ชันสร้างเนื้อหาอีเมลและสร้าง HTML Link (Mailto) ---
def generate_mailto_link(row, alert_type="service"):
    to_addr = str(row['to']) if pd.notna(row['to']) else ""
    cc_addr = str(row['CC']) if pd.notna(row['CC']) else ""
    
    if alert_type == "service":
        subject = f"📢 [แจ้งเตือน] ถึงกำหนดนำรถเข้าศูนย์บริการ: คุณ {row['ชื่อ-นามสกุล']}"
        status_tag = " (⚠️ เกินกำหนดเข้ารับบริการ)" if row['ระยะห่าง'] < 0 else ""
        
        curr_mileage = int(row['เลขไมล์สิ้นสุด']) if pd.notna(row['เลขไมล์สิ้นสุด']) else 0
        next_service = int(row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']) if pd.notna(row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']) else 0
        rem_distance = int(row['ระยะห่าง']) if pd.notna(row['ระยะห่าง']) else 0

        # ใช้ \n สำหรับขึ้นบรรทัดใหม่ในโปรแกรมเมลทั่วไป
        body_text = (
            f"เรียน คุณ {row['ชื่อ-นามสกุล']},\n\n"
            f"ระบบ CMS ตรวจพบว่ารถยนต์ในความดูแลของท่าน ถึงกำหนดต้องเข้ารับการบำรุงรักษา ณ ศูนย์บริการ "
            f"โดยมีรายละเอียดข้อมูลยานพาหนะดังต่อไปนี้ครับ\n\n"
            f"🚗 ข้อมูลยานพาหนะ:\n"
            f"• หมายเลขทะเบียนรถ: {row['ป้ายทะเบียนรถ']}\n"
            f"• เลขไมล์ปัจจุบัน: {curr_mileage:,} กม.\n"
            f"• กำหนดเข้าศูนย์บริการรอบถัดไป: {next_service:,} กม.\n"
            f"• ระยะคงเหลือ: {rem_distance:,} กม. {status_tag}\n\n"
            f"💡 ข้อแนะนำและการดำเนินการ:\n"
            f"1. กรุณาติดต่อและนัดหมายศูนย์บริการล่วงหน้า เพื่อความสะดวกรวดเร็วในการเข้ารับบริการ\n"
            f"2. หากมีข้อสงสัยหรือต้องการข้อมูลเพิ่มเติม กรุณาติดต่อแผนกธุรการ ประจำศูนย์\n\n"
            f"จึงเรียนมาเพื่อทราบและโปรดดำเนินการภายในกำหนดเวลาดังกล่าว\n\n"
            f"ขอแสดงความนับถือ\n"
            f"ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)\n"
            f"จัดทำโดย: แผนกธุรการ POM-NUM_R4"
        )
    else:
        subject = f"📢 [แจ้งเตือน] ถึงกำหนดเปลี่ยนยางรถยนต์: คุณ {row['ชื่อ-นามสกุล']}"
        
        val_limit_months = int(row['limit_months']) if pd.notna(row['limit_months']) else 24
        val_limit_km = int(row['limit_km']) if pd.notna(row['limit_km']) else 50000
        val_current_months = int(row['อายุยาง_เดือน']) if pd.notna(row['อายุยาง_เดือน']) else 0
        val_current_km = int(row['ระยะวิ่งนับจากเปลี่ยนยาง']) if pd.notna(row['ระยะวิ่งนับจากเปลี่ยนยาง']) else 0
        val_last_tyre_km = int(row['เลขไมล์ตอนเปลี่ยนยางล่าสุด']) if pd.notna(row['เลขไมล์ตอนเปลี่ยนยางล่าสุด']) else 0
        val_curr_mileage = int(row['เลขไมล์สิ้นสุด']) if pd.notna(row['เลขไมล์สิ้นสุด']) else 0

        reason_text = f"เนื่องจากใช้ยางครบกำหนด {val_limit_months} เดือน (อายุยางปัจจุบัน: {val_current_months} เดือน)" if row['แจ้งเตือนด้วยเงื่อนไข'] == 'อายุเวลา' else f"เนื่องจากวิ่งครบระยะทาง {val_limit_km:,} กม. (วิ่งไปแล้ว: {val_current_km:,} กม.)"
        date_display = row['วันที่เปลี่ยนยางล่าสุด'].strftime('%d/%m/%Y') if pd.notna(row['วันที่เปลี่ยนยางล่าสุด']) else "-"
        
        body_text = (
            f"เรียน คุณ {row['ชื่อ-นามสกุล']},\n\n"
            f"ระบบ CMS ตรวจพบว่ารถยนต์ในความดูแลของท่าน **ถึงกำหนดต้องเปลี่ยนยางรถยนต์ใหม่** {reason_text} "
            f"เพื่อความปลอดภัยในการขับขี่ โดยมีรายละเอียดดังนี้ครับ\n\n"
            f"🛞 ข้อมูลยางและยานพาหนะ:\n"
            f"• หมายเลขทะเบียนรถ: {row['ป้ายทะเบียนรถ']}\n"
            f"• วันที่เปลี่ยนยางล่าสุด: {date_display} (อายุ {val_current_months} เดือน / กำหนดที่ {val_limit_months} เดือน)\n"
            f"• เลขไมล์ตอนเปลี่ยนยางล่าสุด: {val_last_tyre_km:,} กม.\n"
            f"• เลขไมล์ปัจจุบันล่าสุด: {val_curr_mileage:,} กม.\n"
            f"• ระยะวิ่งรวมของยางชุดนี้: {val_current_km:,} กม. (เกณฑ์จำกัดกำหนดที่ {val_limit_km:,} กม.)\n\n"
            f"💡 ข้อแนะนำและการดำเนินการ:\n"
            f"เพื่อความปลอดภัยในการเดินทาง กรุณาติดต่อเขียนใบเบิกหรือนัดหมายเปลี่ยนยางกับทางแผนกธุรการและฝ่ายจัดซื้อโดยเร็วครับ\n\n"
            f"จึงเรียนมาเพื่อทราบและโปรดดำเนินการภายในกำหนดเวลาดังกล่าว\n\n"
            f"ขอแสดงความนับถือ\n"
            f"ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)\n"
            f"จัดทำโดย: แผนกธุรการ POM-NUM_R4"
        )

    # เข้ารหัสข้อความให้อยู่ในรูปแบบ URL สำหรับใช้ใน mailto link
    encoded_subject = urllib.parse.quote(subject)
    encoded_body = urllib.parse.quote(body_text)
    
    mailto_url = f"mailto:{to_addr}?cc={cc_addr}&subject={encoded_subject}&body={encoded_body}"
    return mailto_url, body_text

# --- 5. เริ่มต้นประมวลผลข้อมูลเมื่ออัปโหลดไฟล์ครบ ---
if file_mileage_ready and file_service_ready and file_tyre_ready and file_config_ready:
    try:
        @st.cache_data(show_spinner=False)
        def process_all_uploaded_data():
            # โหลดไฟล์เงื่อนไขอีเมลจาก buffer
            df_line = pd.read_excel(uploaded_config, sheet_name='LineAPI')
            line_token = str(df_line.iloc[0, 0]).strip() if not df_line.empty else ""
            line_user_id = str(df_line.iloc[0, 1]).strip() if not df_line.empty else ""

            df_e = pd.read_excel(uploaded_config, sheet_name='เงื่อนไข')
            df_e.columns = df_e.columns.str.strip()
            df_e['Name'] = df_e['Name'].astype(str).str.strip()

            # โหลดไฟล์ระยะไมล์การใช้รถล่าสุด
            df_m = pd.read_excel(uploaded_mileage, header=2)
            df_m.columns = df_m.columns.str.strip()
            df_m['ชื่อ-นามสกุล'] = df_m['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_m['ป้ายทะเบียนรถ'] = df_m['ป้ายทะเบียนรถ'].astype(str).str.strip()
            df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
            
            last_m = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('ชื่อ-นามสกุล', as_index=False).last()

            # [ส่วนที่ 1] คำนวณข้อมูลการเข้าศูนย์บริการ
            df_s = pd.read_excel(uploaded_service, header=2)
            df_s.columns = df_s.columns.str.strip()
            df_s['ชื่อพนักงานขับรถปัจจุบัน'] = df_s['ชื่อพนักงานขับรถปัจจุบัน'].astype(str).str.strip()
            df_s['text_next_service'] = df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'].astype(str).str.replace(',', '')
            df_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] = pd.to_numeric(df_s['text_next_service'], errors='coerce')

            combined_s = pd.merge(last_m[['ชื่อ-นามสกุล', 'เลขไมล์สิ้นสุด', 'ป้ายทะเบียนรถ']], 
                                  df_s[['ชื่อพนักงานขับรถปัจจุบัน', 'เลขไมล์เข้าศูนย์บริการรอบถัดไป']], 
                                  how='inner', left_on='ชื่อ-นามสกุล', right_on='ชื่อพนักงานขับรถปัจจุบัน')
            combined_s['ระยะห่าง'] = combined_s['เลขไมล์เข้าศูนย์บริการรอบถัดไป'] - combined_s['เลขไมล์สิ้นสุด']
            service_alerts = pd.merge(combined_s, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')
            service_alerts = service_alerts[service_alerts['ระยะห่าง'] <= 500].sort_values('ระยะห่าง').copy()

            # [ส่วนที่ 2] คำนวณข้อมูลการเปลี่ยนยางรถยนต์
            df_t = pd.read_excel(uploaded_tyre)
            df_t.columns = df_t.columns.str.strip()
            df_t['ชื่อ-นามสกุล'] = df_t['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_t['เลขไมล์ตอนเปลี่ยนยางล่าสุด'] = pd.to_numeric(df_t['เลขไมล์ตอนเปลี่ยนยางล่าสุด'].astype(str).str.replace(',', ''), errors='coerce')
            
            df_t['KM'] = pd.to_numeric(df_t['KM'], errors='coerce').fillna(50000)
            df_t['Months'] = pd.to_numeric(df_t['Months'], errors='coerce').fillna(24)

            def parse_buddhist_date(val):
                if pd.isna(val): return pd.NaT
                try:
                    val_str = str(val).strip().split()[0]
                    if '-' in val_str:
                        parts = val_str.split('-')
                        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                    elif '/' in val_str:
                        parts = val_str.split('/')
                        d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                    else:
                        dt = pd.to_datetime(val, errors='coerce')
                        if pd.isna(dt): return pd.NaT
                        y, m, d = dt.year, dt.month, dt.day
                    if y > 2400: y = y - 543
                    return datetime.datetime(y, m, d)
                except:
                    return pd.NaT

            df_t['วันที่เปลี่ยนยางล่าสุด'] = df_t['วันที่เปลี่ยนยางล่าสุด'].apply(parse_buddhist_date)

            combined_t = pd.merge(last_m[['ชื่อ-นามสกุล', 'เลขไมล์สิ้นสุด', 'ป้ายทะเบียนรถ']], 
                                  df_t[['ชื่อ-นามสกุล', 'วันที่เปลี่ยนยางล่าสุด', 'เลขไมล์ตอนเปลี่ยนยางล่าสุด', 'KM', 'Months']], 
                                  on='ชื่อ-นามสกุล', how='inner')
            combined_t['ระยะวิ่งนับจากเปลี่ยนยาง'] = combined_t['เลขไมล์สิ้นสุด'] - combined_t['เลขไมล์ตอนเปลี่ยนยางล่าสุด']
            
            current_date = datetime.datetime.now()
            combined_t['อายุยาง_เดือน'] = ((current_date.year - combined_t['วันที่เปลี่ยนยางล่าสุด'].dt.year) * 12 + 
                                          (current_date.month - combined_t['วันที่เปลี่ยนยางล่าสุด'].dt.month))
            
            combined_t['limit_km'] = combined_t['KM']
            combined_t['limit_months'] = combined_t['Months']
            combined_t['alert_km_trigger'] = combined_t['KM'] - 5000
            combined_t['alert_month_trigger'] = combined_t['Months'] - 1

            tyre_alerts = combined_t[(combined_t['ระยะวิ่งนับจากเปลี่ยนยาง'] >= combined_t['alert_km_trigger']) | 
                                     (combined_t['อายุยาง_เดือน'] >= combined_t['alert_month_trigger'])].copy()
            
            tyre_alerts['แจ้งเตือนด้วยเงื่อนไข'] = tyre_alerts.apply(
                lambda r: 'ระยะทาง' if r['ระยะวิ่งนับจากเปลี่ยนยาง'] >= r['alert_km_trigger'] else 'อายุเวลา', axis=1
            )
            tyre_alerts = pd.merge(tyre_alerts, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')

            return service_alerts, tyre_alerts, line_token, line_user_id

        with st.status("🚀 กำลังวิเคราะห์และคำนวณข้อมูลจากไฟล์ที่อัปโหลด...", expanded=False) as status:
            service_alerts, tyre_alerts, line_token, line_user_id = process_all_uploaded_data()
            status.update(label="✅ คำนวณข้อมูลรถและยางเสร็จสิ้นเรียบร้อย!", state="complete")

        # --- 6. แสดงผลแบ่งแท็บแยกประเภทเพื่อความเรียบร้อยบนหน้าเว็บ ---
        tab1, tab2 = st.tabs(["🚗 รายการเข้าศูนย์บริการ", "🛞 รายการเปลี่ยนยางรถยนต์"])

        # --- TAB 1: บำรุงรักษาเข้าศูนย์ ---
        with tab1:
            if not service_alerts.empty:
                st.subheader("📬 รายการแจ้งเตือนการเข้าศูนย์ (คลิกลิงก์เพื่อส่งเมล)")
                
                for index, row in service_alerts.iterrows():
                    mailto_link, body_preview = generate_mailto_link(row, alert_type="service")
                    
                    with st.container():
                        c_info, c_action = st.columns([7, 3])
                        with c_info:
                            status_color = 'red' if row['ระยะห่าง'] < 0 else '#d97706'
                            val_curr_m = int(row['เลขไมล์สิ้นสุด']) if pd.notna(row['เลขไมล์สิ้นสุด']) else 0
                            val_rem_d = int(row['ระยะห่าง']) if pd.notna(row['ระยะห่าง']) else 0
                            st.markdown(f'<div class="card-container"><strong>👤 คุณ {row["ชื่อ-นามสกุล"]}</strong> | ทะเบียน: <b>{row["ป้ายทะเบียนรถ"]}</b><br>เลขไมล์ปัจจุบัน: {val_curr_m:,} กม. | ระยะคงเหลือ: <span style="color:{status_color}; font-weight:bold;">{val_rem_d:,} กม.</span></div>', unsafe_allow_html=True)
                        with c_action:
                            # ใช้ st.link_button เพื่อเปิดโปรแกรมเมลในเครื่องทันที
                            st.link_button("📧 กดเพื่อส่งเมลด้วย Outlook", mailto_link, type="primary", use_container_width=True)
                            with st.expander("🔍 ดูข้อความ"):
                                st.text(body_preview)
            else:
                st.success("✨ ไม่มีรถคันไหนถึงกำหนดเข้าศูนย์บริการในไฟล์ชุดนี้")

        # --- TAB 2: เปลี่ยนยางรถยนต์ ---
        with tab2:
            if not tyre_alerts.empty:
                st.subheader("📬 รายการแจ้งเตือนเปลี่ยนยาง (คลิกลิงก์เพื่อส่งเมล)")

                for index, row in tyre_alerts.iterrows():
                    mailto_link, body_preview = generate_mailto_link(row, alert_type="tyre")
                    
                    with st.container():
                        c_info, c_action = st.columns([7, 3])
                        with c_info:
                            v_limit_k = int(row['limit_km']) if pd.notna(row['limit_km']) else 50000
                            v_limit_m = int(row['limit_months']) if pd.notna(row['limit_months']) else 24
                            v_curr_m = int(row['ระยะวิ่งนับจากเปลี่ยนยาง']) if pd.notna(row['ระยะวิ่งนับจากเปลี่ยนยาง']) else 0
                            v_age_m = int(row['อายุยาง_เดือน']) if pd.notna(row['อายุยาง_เดือน']) else 0

                            badge = f"🔵 เตือนด้วยอายุยาง (กำหนด {v_limit_m} เดือน)" if row['แจ้งเตือนด้วยเงื่อนไข'] == 'อายุเวลา' else f"🟠 เตือนด้วยระยะทาง (กำหนด {v_limit_k:,} กม.)"
                            st.markdown(f'<div class="card-container"><strong>👤 คุณ {row["ชื่อ-นามสกุล"]}</strong> | ทะเบียน: <b>{row["ป้ายทะเบียนรถ"]}</b><br>วิ่งไปแล้ว: {v_curr_m:,} กม. / เกณฑ์ {v_limit_k:,} กม. | อายุยาง: {v_age_m} เดือน / เกณฑ์ {v_limit_m} เดือน | <span style="color:#2563eb; font-weight:bold;">{badge}</span></div>', unsafe_allow_html=True)
                        with c_action:
                            st.link_button("📧 กดเพื่อส่งเมลด้วย Outlook", mailto_link, type="primary", use_container_width=True)
                            with st.expander("🔍 ดูข้อความ"):
                                st.text(body_preview)
            else:
                st.success("✨ ไม่มีรถคันไหนเข้าเกณฑ์ต้องเปลี่ยนยางรถยนต์ในขณะนี้")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดทางเทคนิคในการอ่านไฟล์: {e}")
else:
    st.info("💡 **คำแนะนำ:** กรุณานำไฟล์ Excel ทั้ง 4 ไฟล์มาลากวางอัปโหลดที่เมนูด้านซ้าย (Sidebar) ให้ครบถ้วน ระบบจึงจะเริ่มประมวลผลข้อมูลให้ครับ")

# --- Footer เครดิต ---
st.markdown('<div class="footer">Developed by <b>ITsupportR4</b> | CMS v3.0 (Cloud Web Version)</div>', unsafe_allow_html=True)
