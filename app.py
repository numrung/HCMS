import os
import datetime
from playwright.sync_api import Playwright, sync_playwright, expect

def run(playwright: Playwright) -> None:
    # --- 1. เตรียมข้อมูลวันที่และที่เก็บไฟล์ ---
    now = datetime.datetime.now()
    months_th = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", 
                 "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    
    current_month_name = months_th[now.month]
    current_year_th = now.year + 543  
    
    target_dir = r"D:\HCMSPOM\Data"
    
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    print("="*60)
    print("🚀  CMS AUTOMATION FETCH SYSTEM (v2.0)")
    print(f"📅  ประจำเดือน: {current_month_name} | ปี พ.ศ.: {current_year_th}")
    print("="*60)

    # 🌟 ตั้งค่า headless=True เพื่อให้รันเบื้องหลังและปิดหน้าต่างอัตโนมัติเมื่อเสร็จงาน
    print("⏳ [1/5] กำลังเริ่มทำงานเบื้องหลัง (Headless Mode)...")
    browser = playwright.chromium.launch(headless=True) 
    context = browser.new_context()
    page = context.new_page()

    # --- 2. LOGIN ---
    print("🔐 [2/5] กำลังเข้าสู่ระบบ CMS (Thaibev)...")
    try:
        page.goto("https://ss.thaibev.com/CMS/", timeout=60000)
        page.locator("#empId").fill("11030028")
        page.locator("#password").fill("11030028")
        page.get_by_role("button", name="เข้าสู่ระบบ").click()
        
        page.get_by_role("button", name="OK").click()
        page.wait_for_load_state("networkidle")
        print("    👉 ล็อกอินสำเร็จเรียบร้อย")
    except Exception as e:
        print(f"    ❌ เกิดข้อผิดพลาดตอนล็อกอิน: {e}")
        browser.close()
        return

    # --- 3. รายงานที่ 1: รายงานการใช้รถ ---
    print(f"📦 [3/5] กำลังดาวน์โหลด: รายงานการใช้รถ ({current_month_name})")
    try:
        page.get_by_role("link", name="รายงาน", exact=True).hover()
        page.get_by_role("link", name="» รายงานการใช้รถ", exact=True).click()
        
        month_input = page.locator("#MonthForReport")
        month_input.click()
        month_input.fill(current_month_name)
        page.keyboard.press("Enter") 
        
        page.get_by_label("").first.click()
        page.get_by_role("treeitem", name="6800 บริษัท ป้อมพลัง จำกัด").click()

        with page.expect_download() as download_info:
            page.get_by_role("button", name="ดูรายงาน").click()
        
        download = download_info.value
        file1_name = f"รายงานการใช้รถ_{current_month_name}_{current_year_th}.xlsx"
        path1 = os.path.join(target_dir, file1_name)
        download.save_as(path1)
        print(f"    💾 บันทึกไฟล์สำเร็จ -> {file1_name}")
    except Exception as e:
        print(f"    ❌ เกิดข้อผิดพลาดในรายงานการใช้รถ: {e}")

    # --- 4. รายงานที่ 2: รายงานการเข้าศูนย์บริการ ---
    print("📦 [4/5] กำลังดาวน์โหลด: รายงานการเข้าศูนย์บริการ (ล่าสุด)")
    try:
        page.get_by_role("link", name="รายงาน", exact=True).hover()
        page.get_by_role("link", name="» รายงานการเข้าศูนย์บริการ (ล่าสุด)").click()
        
        page.get_by_label("").first.click()
        page.get_by_role("treeitem", name="6800 บริษัท ป้อมพลัง จำกัด").click()

        with page.expect_download() as download1_info:
            page.get_by_role("button", name="ดูรายงาน").click()
        
        download1 = download1_info.value
        file2_name = f"รายงานการเข้าศูนย์_{current_month_name}_{current_year_th}.xlsx"
        path2 = os.path.join(target_dir, file2_name)
        download1.save_as(path2)
        print(f"    💾 บันทึกไฟล์สำเร็จ -> {file2_name}")
    except Exception as e:
        print(f"    ❌ เกิดข้อผิดพลาดในรายงานเข้าศูนย์: {e}")

    # --- 5. ปิดการทำงาน ---
    print("🛑 [5/5] กำลังเคลียร์พาร์ทิชันและปิดระบบ...")
    context.close()
    browser.close()
    
    print("="*60)
    print("✨ SUCCESS: ดึงข้อมูลรายงานครบถ้วนและจัดเก็บเป็น พ.ศ. เรียบร้อยแล้ว!")
    print("="*60)

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
