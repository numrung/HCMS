import os
import json
import datetime
import urllib.parse
import requests
import pandas as pd
import streamlit as st

# นำเข้าโมดูลสำหรับควบคุมโปรแกรม Outlook ในเครื่อง Windows
try:
    import win32com.client as win32
except ImportError:
    win32 = None

# --- 1. ตั้งค่าหน้าจอ & Theme ---
st.set_page_config(page_title="CMS Maintenance System", layout="wide", page_icon="🚗")

# Custom CSS สำหรับความสวยงามและควบคุมโครงสร้างหน้าเว็บ
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
st.caption("ระบบดึงไฟล์อัตโนมัติ ตรวจสอบระยะเข้าศูนย์ และอายุยางรถยนต์ เลือกส่งอีเมลแยกตามหมวดหมู่ผ่าน Outlook")

# --- 2. ส่วนคำนวณหาชื่อไฟล์ประจำเดือนปัจจุบันอัตโนมัติ (ปรับเป็น พ.ศ. ให้ค้นหาในเครื่องเจอ) ---
now = datetime.datetime.now()
months_th = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", 
             "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
current_month_name = months_th[now.month]
current_year_th = now.year + 543  # ปรับให้หาไฟล์เวอร์ชัน พ.ศ. ตามชื่อไฟล์จริงในเครื่อง

# กำหนด Path ที่อยู่ของไฟล์ในเครื่องคอมพิวเตอร์ของคุณ
data_dir = r"D:\HCMSPOM\Data"
config_file_path = r"D:\HCMSPOM\logic\Email.xlsx"

# Path ไฟล์ยางตามที่คุณระบุ
tyre_file_path = r"D:\HCMSPOM\Service\รายงานการเปลี่ยนยาง.xlsx"

# สร้างรูปแบบชื่อไฟล์ที่ระบบต้องตามหา
expected_mileage_file = f"รายงานการใช้รถ_{current_month_name}_{current_year_th}.xlsx"
expected_service_file = f"รายงานการเข้าศูนย์_{current_month_name}_{current_year_th}.xlsx"

path_mileage = os.path.join(data_dir, expected_mileage_file)
path_service = os.path.join(data_dir, expected_service_file)

# --- 3. Sidebar แสดงสถานะการตรวจพบไฟล์ในเครื่อง ---
with st.sidebar:
    st.header("📖 Status สถานะไฟล์")
    
    if os.path.exists(path_mileage):
        st.success(f"🟢 พบไฟล์รายงานการใช้รถ:\n`{expected_mileage_file}`")
        file_mileage_ready = True
    else:
        st.error(f"🔴 ไม่พบไฟล์ประจำเดือนนี้:\n`{expected_mileage_file}`")
        file_mileage_ready = False

    if os.path.exists(path_service):
        st.success(f"🟢 พบไฟล์ข้อมูลเข้าศูนย์:\n`{expected_service_file}`")
        file_service_ready = True
    else:
        st.error(f"🔴 ไม่พบไฟล์ประจำเดือนนี้:\n`{expected_service_file}`")
        file_service_ready = False

    # ตรวจสอบไฟล์ยางที่โฟลเดอร์ใหม่
    if os.path.exists(tyre_file_path):
        st.success(f"🟢 พบไฟล์ประวัติการเปลี่ยนยางที่:\n`D:\\HCMSPOM\\Service\\รายงานการเปลี่ยนยาง.xlsx`")
        file_tyre_ready = True
    else:
        st.error(f"🔴 ไม่พบไฟล์ประวัติยางที่:\n`D:\\HCMSPOM\\Service\\รายงานการเปลี่ยนยาง.xlsx`")
        file_tyre_ready = False

    if os.path.exists(config_file_path):
        st.success(f"🟢 พบไฟล์เงื่อนไข & API:\n`Email.xlsx`")
        file_config_ready = True
    else:
        st.error(f"🔴 ไม่พบไฟล์ตั้งค่าที่:\n`D:\\HCMSPOM\\logic\\Email.xlsx`")
        file_config_ready = False
        
    st.divider()
    st.write("💻 **Developer:** ITsupportR4")

# --- 4. ฟังก์ชันสร้างเนื้อหาอีเมล ---
def get_mail_content(row, alert_type="service"):
    to_addr = str(row['to']) if pd.notna(row['to']) else ""
    cc_addr = str(row['CC']) if pd.notna(row['CC']) else ""
    
    if alert_type == "service":
        subject = f"📢 [แจ้งเตือน] ถึงกำหนดนำรถเข้าศูนย์บริการ: คุณ {row['ชื่อ-นามสกุล']}"
        status_tag = " <span style='color: red; font-weight: bold;'>(⚠️ เกินกำหนดเข้ารับบริการ)</span>" if row['ระยะห่าง'] < 0 else ""
        
        # ป้องกันเลขไมล์มีทศนิยมตอนแปลงข้อมูลเข้า HTML
        curr_mileage = int(row['เลขไมล์สิ้นสุด']) if pd.notna(row['เลขไมล์สิ้นสุด']) else 0
        next_service = int(row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']) if pd.notna(row['เลขไมล์เข้าศูนย์บริการรอบถัดไป']) else 0
        rem_distance = int(row['ระยะห่าง']) if pd.notna(row['ระยะห่าง']) else 0

        html_body = f"""
        <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 16px; color: #333333; line-height: 1.6;">
            <p>เรียน คุณ {row['ชื่อ-นามสกุล']},</p>
            <p>ระบบ CMS ตรวจพบว่ารถยนต์ในความดูแลของท่าน ถึงกำหนดต้องเข้ารับการบำรุงรักษา ณ ศูนย์บริการ โดยมีรายละเอียดข้อมูลยานพาหนะดังต่อไปนี้ครับ</p>
            <h3 style="color: #1e3a8a; font-size: 18px; margin-top: 22px; margin-bottom: 12px;">🚗 ข้อมูลยานพาหนะ:</h3>
            <table style="border-collapse: collapse; width: 100%; max-width: 580px; font-size: 16px; margin-bottom: 22px; border: 1px solid #e5e7eb;">
                <thead>
                    <tr style="background-color: #1e3a8a; color: white;">
                        <th style="border: 1px solid #1e3a8a; text-align: left; padding: 12px; width: 50%; font-weight: bold;">รายการข้อมูล</th>
                        <th style="border: 1px solid #1e3a8a; text-align: left; padding: 12px; width: 50%; font-weight: bold;">รายละเอียดรถยนต์</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• หมายเลขทะเบียนรถ</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{row['ป้ายทะเบียนรถ']}</td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• เลขไมล์ปัจจุบัน</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{curr_mileage:,} กม.</td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• กำหนดเข้าศูนย์บริการรอบถัดไป</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{next_service:,} กม.</td>
                    </tr>
                    <tr style="background-color: #fef2f2;">
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #fee2e2;">• ระยะคงเหลือ</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; color: {'#dc2626' if rem_distance < 0 else '#d97706'};">{rem_distance:,} กม.{status_tag}</td>
                    </tr>
                </tbody>
            </table>
            <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px; margin-top: 20px; max-width: 580px; border-radius: 0 4px 4px 0; font-size: 15px;">
                <strong style="color: #b45309; font-size: 16px;">💡 ข้อแนะนำและการดำเนินการ:</strong><br>
                1. กรุณาติดต่อและนัดหมายศูนย์บริการล่วงหน้า เพื่อความสะดวกรวดเร็วในการเข้ารับบริการ<br>
                2. หากมีข้อสงสัยหรือต้องการข้อมูลเพิ่มเติม กรุณาติดต่อแผนกธุรการ ประจำศูนย์
            </div>
            <p style="margin-top: 25px;">จึงเรียนมาเพื่อทราบและโปรดดำเนินการภายในกำหนดเวลาดังกล่าว</p>
            <p>ขอแสดงความนับถือ</p>
            <p style="line-height: 1.4;">
                <strong>ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)</strong><br>
                <span style="color: #6b7280; font-size: 14px;">จัดทำโดย: แผนกธุรการ POM-NUM_R4</span>
            </p>
        </body>
        </html>
        """
    else:
        subject = f"📢 [แจ้งเตือน] ถึงกำหนดเปลี่ยนยางรถยนต์: คุณ {row['ชื่อ-นามสกุล']}"
        
        # ป้องกันประเภทข้อมูลหลุดเป็นทศนิยม (.0) หลังจาก Merge ตารางเงื่อนไข
        val_limit_months = int(row['limit_months']) if pd.notna(row['limit_months']) else 24
        val_limit_km = int(row['limit_km']) if pd.notna(row['limit_km']) else 50000
        val_current_months = int(row['อายุยาง_เดือน']) if pd.notna(row['อายุยาง_เดือน']) else 0
        val_current_km = int(row['ระยะวิ่งนับจากเปลี่ยนยาง']) if pd.notna(row['ระยะวิ่งนับจากเปลี่ยนยาง']) else 0
        val_last_tyre_km = int(row['เลขไมล์ตอนเปลี่ยนยางล่าสุด']) if pd.notna(row['เลขไมล์ตอนเปลี่ยนยางล่าสุด']) else 0
        val_curr_mileage = int(row['เลขไมล์สิ้นสุด']) if pd.notna(row['เลขไมล์สิ้นสุด']) else 0

        reason_text = f"เนื่องจากใช้ยางครบกำหนด {val_limit_months} เดือน (อายุยางปัจจุบัน: {val_current_months} เดือน)" if row['แจ้งเตือนด้วยเงื่อนไข'] == 'อายุเวลา' else f"เนื่องจากวิ่งครบระยะทาง {val_limit_km:,} กม. (วิ่งไปแล้ว: {val_current_km:,} กม.)"
        date_display = row['วันที่เปลี่ยนยางล่าสุด'].strftime('%d/%m/%Y') if pd.notna(row['วันที่เปลี่ยนยางล่าสุด']) else "-"
        
        html_body = f"""
        <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 16px; color: #333333; line-height: 1.6;">
            <p>เรียน คุณ {row['ชื่อ-นามสกุล']},</p>
            <p>ระบบ CMS ตรวจพบว่ารถยนต์ในความดูแลของท่าน **ถึงกำหนดต้องเปลี่ยนยางรถยนต์ใหม่** {reason_text} เพื่อความปลอดภัยในการขับขี่ โดยมีรายละเอียดดังนี้ครับ</p>
            <h3 style="color: #1e3a8a; font-size: 18px; margin-top: 22px; margin-bottom: 12px;">🛞 ข้อมูลยางและยานพาหนะ:</h3>
            <table style="border-collapse: collapse; width: 100%; max-width: 580px; font-size: 16px; margin-bottom: 22px; border: 1px solid #e5e7eb;">
                <thead>
                    <tr style="background-color: #2563eb; color: white;">
                        <th style="border: 1px solid #e5e7eb; text-align: left; padding: 12px; width: 50%; font-weight: bold;">รายการข้อมูล</th>
                        <th style="border: 1px solid #e5e7eb; text-align: left; padding: 12px; width: 50%; font-weight: bold;">รายละเอียดการเปลี่ยนยาง</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• หมายเลขทะเบียนรถ</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{row['ป้ายทะเบียนรถ']}</td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• วันที่เปลี่ยนยางล่าสุด</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{date_display} (อายุ {val_current_months} เดือน / กำหนดที่ {val_limit_months} เดือน)</td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• เลขไมล์ตอนเปลี่ยนยางล่าสุด</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{val_last_tyre_km:,} กม.</td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #f9fafb;">• เลขไมล์ปัจจุบันล่าสุด</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px;">{val_curr_mileage:,} กม.</td>
                    </tr>
                    <tr style="background-color: #fef2f2;">
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; background-color: #fee2e2;">• ระยะวิ่งรวมของยางชุดนี้</td>
                        <td style="border: 1px solid #e5e7eb; padding: 12px; font-weight: bold; color: #dc2626;">{val_current_km:,} กม. (เกณฑ์จำกัดกำหนดที่ {val_limit_km:,} กม.)</td>
                    </tr>
                </tbody>
            </table>
            <div style="background-color: #eff6ff; border-left: 4px solid #2563eb; padding: 15px; margin-top: 20px; max-width: 580px; border-radius: 0 4px 4px 0; font-size: 15px;">
                <strong style="color: #1e40af; font-size: 16px;">💡 ข้อแนะนำและการดำเนินการ:</strong><br>
                เพื่อความปลอดภัยในการเดินทาง กรุณาติดต่อเขียนใบเบิกหรือนัดหมายเปลี่ยนยางกับทางแผนกธุรการและฝ่ายจัดซื้อโดยเร็วครับ
            </div>
            <p style="margin-top: 25px;">จึงเรียนมาเพื่อทราบและโปรดดำเนินการภายในกำหนดเวลาดังกล่าว</p>
            <p>ขอแสดงความนับถือ</p>
            <p style="line-height: 1.4;">
                <strong>ระบบแจ้งเตือนการบำรุงรักษารถยนต์ (CMS)</strong><br>
                <span style="color: #6b7280; font-size: 14px;">จัดทำโดย: แผนกธุรการ POM-NUM_R4</span>
            </p>
        </body>
        </html>
        """
    return to_addr, cc_addr, subject, html_body

# --- 5. ฟังก์ชันสั่งงานโปรแกรม Outlook ---
def send_via_local_outlook(to_addr, cc_addr, subject, html_body):
    if win32 is None:
        st.error("❌ ไม่สามารถรันระบบควบคุม Outlook ได้เนื่องจากขาดโมดูล `pywin32`")
        return False
    try:
        import pythoncom
        pythoncom.CoInitialize()
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        mail.To = to_addr
        if cc_addr and cc_addr.strip() != "":
            mail.CC = cc_addr
        mail.Subject = subject
        mail.HTMLBody = html_body
        
        mail.Send()
        return True
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการส่งหาคุณ {to_addr}: {e}")
        return False

# --- 6. เริ่มต้นประมวลผลข้อมูลทั้งหมด ---
if file_mileage_ready and file_service_ready and file_tyre_ready and file_config_ready:
    try:
        @st.cache_data(show_spinner=False)
        def process_all_data():
            # โหลดไฟล์เงื่อนไขอีเมล
            df_line = pd.read_excel(config_file_path, sheet_name='LineAPI')
            line_token = str(df_line.iloc[0, 0]).strip() if not df_line.empty else ""
            line_user_id = str(df_line.iloc[0, 1]).strip() if not df_line.empty else ""

            df_e = pd.read_excel(config_file_path, sheet_name='เงื่อนไข')
            df_e.columns = df_e.columns.str.strip()
            df_e['Name'] = df_e['Name'].astype(str).str.strip()

            # โหลดไฟล์ระยะไมล์การใช้รถล่าสุด
            df_m = pd.read_excel(path_mileage, header=2)
            df_m.columns = df_m.columns.str.strip()
            df_m['ชื่อ-นามสกุล'] = df_m['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_m['ป้ายทะเบียนรถ'] = df_m['ป้ายทะเบียนรถ'].astype(str).str.strip()
            df_m['เลขไมล์สิ้นสุด'] = pd.to_numeric(df_m['เลขไมล์สิ้นสุด'].astype(str).str.replace(',', ''), errors='coerce')
            
            last_m = df_m.dropna(subset=['เลขไมล์สิ้นสุด']).groupby('ชื่อ-นามสกุล', as_index=False).last()

            # [ส่วนที่ 1] คำนวณข้อมูลการเข้าศูนย์บริการ
            df_s = pd.read_excel(path_service, header=2)
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

            # [ส่วนที่ 2] คำนวณข้อมูลการเปลี่ยนยางรถยนต์ (รวมคอลัมน์ KM, Months แบบ Dynamic Threshold)
            df_t = pd.read_excel(tyre_file_path)
            df_t.columns = df_t.columns.str.strip()
            df_t['ชื่อ-นามสกุล'] = df_t['ชื่อ-นามสกุล'].astype(str).str.strip()
            df_t['เลขไมล์ตอนเปลี่ยนยางล่าสุด'] = pd.to_numeric(df_t['เลขไมล์ตอนเปลี่ยนยางล่าสุด'].astype(str).str.replace(',', ''), errors='coerce')
            
            # ดึงข้อมูลกำหนดระยะของแต่ละบุคคล (ถ้าไม่ได้กรอกให้สลับไปใช้เกณฑ์ Standard ทันทีด้วย .fillna())
            df_t['KM'] = pd.to_numeric(df_t['KM'], errors='coerce').fillna(50000)
            df_t['Months'] = pd.to_numeric(df_t['Months'], errors='coerce').fillna(24)

            # ฟังก์ชันแปลงวันที่แบบปลอดภัย (พ.ศ. -> ค.ศ.)
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
            
            # คำนวณอายุยางเป็นจำนวนเดือน
            current_date = datetime.datetime.now()
            combined_t['อายุยาง_เดือน'] = ((current_date.year - combined_t['วันที่เปลี่ยนยางล่าสุด'].dt.year) * 12 + 
                                          (current_date.month - combined_t['วันที่เปลี่ยนยางล่าสุด'].dt.month))
            
            # กำหนดช่วงเกณฑ์การเตือนล่วงหน้าอ้างอิงตามเลขแต่ละคน (เตือนล่วงหน้า 5,000 กม. หรือ 1 เดือน)
            combined_t['limit_km'] = combined_t['KM']
            combined_t['limit_months'] = combined_t['Months']
            combined_t['alert_km_trigger'] = combined_t['KM'] - 5000
            combined_t['alert_month_trigger'] = combined_t['Months'] - 1

            tyre_alerts = combined_t[(combined_t['ระยะวิ่งนับจากเปลี่ยนยาง'] >= combined_t['alert_km_trigger']) | 
                                     (combined_t['อายุยาง_เดือน'] >= combined_t['alert_month_trigger'])].copy()
            
            tyre_alerts['แจ้งเตือนด้วยเงื่อนไข'] = tyre_alerts.apply(
                lambda r: 'ระยะทาง' if r['ระยะวิ่งนับจากเปลี่ยนยาง'] >= r['alert_km_trigger'] else 'อายุเวลา', axis=1
            )
            # ดึงอีเมลผู้รับและ CC ประจำตัวของคนขับรถมาจากหน้าต่าง df_e
            tyre_alerts = pd.merge(tyre_alerts, df_e, left_on='ชื่อ-นามสกุล', right_on='Name', how='left')

            return service_alerts, tyre_alerts, line_token, line_user_id

        with st.status("🚀 กำลังวิเคราะห์และคำนวณข้อมูลแบบไฮบริด...", expanded=False) as status:
            service_alerts, tyre_alerts, line_token, line_user_id = process_all_data()
            status.update(label="✅ คำนวณข้อมูลรถและยางเสร็จสิ้นเรียบร้อย!", state="complete")

        # --- 7. แสดงผลแบ่งแท็บแยกประเภทเพื่อความเรียบร้อยบนหน้าเว็บ ---
        tab1, tab2 = st.tabs(["🚗 รายการเข้าศูนย์บริการ", "🛞 รายการเปลี่ยนยางรถยนต์"])

        # --- TAB 1: บำรุงรักษาเข้าศูนย์ ---
        with tab1:
            if not service_alerts.empty:
                st.subheader("📬 เลือกส่งเมลแจ้งเตือนการเข้าศูนย์")
                
                col_s1, col_s2, _ = st.columns([1.5, 1.5, 7])
                with col_s1:
                    if st.button("✅ เลือกทุกคน (ศูนย์)", key="all_s"):
                        for i in service_alerts.index: st.session_state[f"chks_{i}"] = True
                with col_s2:
                    if st.button("❌ ล้างทั้งหมด (ศูนย์)", key="clr_s"):
                        for i in service_alerts.index: st.session_state[f"chks_{i}"] = False

                selected_s = []
                for index, row in service_alerts.iterrows():
                    to_addr, cc_addr, subject, html_body = get_mail_content(row, alert_type="service")
                    chk_key = f"chks_{index}"
                    if chk_key not in st.session_state: st.session_state[chk_key] = True
                    
                    with st.container():
                        c_check, c_info, c_preview = st.columns([0.5, 6.5, 3])
                        with c_check:
                            if st.checkbox("เลือก", key=chk_key, label_visibility="collapsed"): selected_s.append(row)
                        with c_info:
                            status_color = 'red' if row['ระยะห่าง'] < 0 else '#d97706'
                            val_curr_m = int(row['เลขไมล์สิ้นสุด']) if pd.notna(row['เลขไมล์สิ้นสุด']) else 0
                            val_rem_d = int(row['ระยะห่าง']) if pd.notna(row['ระยะห่าง']) else 0
                            st.markdown(f'<div class="card-container"><strong>👤 คุณ {row["ชื่อ-นามสกุล"]}</strong> | ทะเบียน: <b>{row["ป้ายทะเบียนรถ"]}</b><br>เลขไมล์ปัจจุบัน: {val_curr_m:,} กม. | ระยะคงเหลือ: <span style="color:{status_color}; font-weight:bold;">{val_rem_d:,} กม.</span></div>', unsafe_allow_html=True)
                        with c_preview:
                            with st.expander("🔍 ดูข้อความ"): st.components.v1.html(html_body, height=250, scrolling=True)

                if st.button(f"🚀 ส่งเมลแจ้งเตือนเข้าศูนย์ ({len(selected_s)} คน)", type="primary", use_container_width=True, disabled=len(selected_s)==0):
                    sc = sum([1 for r in selected_s if send_via_local_outlook(*get_mail_content(r, "service"))])
                    st.success(f"🎉 สั่ง Outlook ส่งเมลแจ้งเตือนเข้าศูนย์สำเร็จ {sc} คน")
            else:
                st.success("✨ ไม่มีรถคันไหนถึงกำหนดเข้าศูนย์บริการในเดือนนี้")

        # --- TAB 2: เปลี่ยนยางรถยนต์ ---
        with tab2:
            if not tyre_alerts.empty:
                st.subheader("📬 เลือกส่งเมลแจ้งเตือนเปลี่ยนยาง")
                
                col_t1, col_t2, _ = st.columns([1.5, 1.5, 7])
                with col_t1:
                    if st.button("✅ เลือกทุกคน (ยาง)", key="all_t"):
                        for i in tyre_alerts.index: st.session_state[f"chkt_{i}"] = True
                with col_t2:
                    if st.button("❌ ล้างทั้งหมด (ยาง)", key="clr_t"):
                        for i in tyre_alerts.index: st.session_state[f"chkt_{i}"] = False

                selected_t = []
                for index, row in tyre_alerts.iterrows():
                    to_addr, cc_addr, subject, html_body = get_mail_content(row, alert_type="tyre")
                    chk_key = f"chkt_{index}"
                    if chk_key not in st.session_state: st.session_state[chk_key] = True
                    
                    with st.container():
                        c_check, c_info, c_preview = st.columns([0.5, 6.5, 3])
                        with c_check:
                            if st.checkbox("เลือก", key=chk_key, label_visibility="collapsed"): selected_t.append(row)
                        with c_info:
                            # ป้องกันเลขทศนิยมโผล่บนหน้าเว็บ UI การ์ดแสดงผล
                            v_limit_m = int(row['limit_months']) if pd.notna(row['limit_months']) else 24
                            v_limit_k = int(row['limit_km']) if pd.notna(row['limit_km']) else 50000
                            v_curr_m = int(row['ระยะวิ่งนับจากเปลี่ยนยาง']) if pd.notna(row['ระยะวิ่งนับจากเปลี่ยนยาง']) else 0
                            v_age_m = int(row['อายุยาง_เดือน']) if pd.notna(row['อายุยาง_เดือน']) else 0

                            badge = f"🔵 เตือนด้วยอายุยาง (กำหนด {v_limit_m} เดือน)" if row['แจ้งเตือนด้วยเงื่อนไข'] == 'อายุเวลา' else f"🟠 เตือนด้วยระยะทาง (กำหนด {v_limit_k:,} กม.)"
                            st.markdown(f'<div class="card-container"><strong>👤 คุณ {row["ชื่อ-นามสกุล"]}</strong> | ทะเบียน: <b>{row["ป้ายทะเบียนรถ"]}</b><br>วิ่งไปแล้ว: {v_curr_m:,} กม. / เกณฑ์ {v_limit_k:,} กม. | อายุยาง: {v_age_m} เดือน / เกณฑ์ {v_limit_m} เดือน | <span style="color:#2563eb; font-weight:bold;">{badge}</span></div>', unsafe_allow_html=True)
                        with c_preview:
                            with st.expander("🔍 ดูข้อความ"): st.components.v1.html(html_body, height=250, scrolling=True)

                if st.button(f"🚀 ส่งเมลแจ้งเตือนเปลี่ยนยาง ({len(selected_t)} คน)", type="primary", use_container_width=True, disabled=len(selected_t)==0):
                    tc = sum([1 for r in selected_t if send_via_local_outlook(*get_mail_content(r, "tyre"))])
                    st.success(f"🎉 สั่ง Outlook ส่งเมลแจ้งเตือนเปลี่ยนยางสำเร็จ {tc} คน")
            else:
                st.success("✨ ไม่มีรถคันไหนเข้าเกณฑ์ต้องเปลี่ยนยางรถยนต์ในขณะนี้")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดทางเทคนิค: {e}")
else:
    st.info("💡 **คำแนะนำ:** กรุณาเตรียมไฟล์ข้อมูลให้ครบถ้วนตามโฟลเดอร์ที่ระบบกำหนด ระบบจึงจะเริ่มประมวลผลไฮบริดครับ")

# --- Footer เครดิต ---
st.markdown('<div class="footer">Developed by <b>ITsupportR4</b> | CMS v2.0 (Hybrid Maintenance & Tyre System)</div>', unsafe_allow_html=True)
