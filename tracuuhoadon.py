# Import thư viện cần thiết
import os
import time
import shutil
import xml.etree.ElementTree as ET
import pandas as pd
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from openpyxl import load_workbook, Workbook

# Mở trình duyệt với cấu hình thư mục tải về
def open_browser(thu_muc_tai_hoa_don):
    os.makedirs(thu_muc_tai_hoa_don, exist_ok=True)
    options = Options()
    options.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True, 
        "download.default_directory": thu_muc_tai_hoa_don,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    })
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    service = Service()
    driver = webdriver.Chrome(service=service, options=options)
    return driver, WebDriverWait(driver, 10)

# Tra cứu hóa đơn trên từng trang web tương ứng
def tra_cuu_hoa_don(driver, wait, ma_so_thue, ma_tra_cuu, url):
    try:    
        driver.get(url)
        
        # Tra cứu trên FPT
        if "https://tracuuhoadon.fpt.com.vn/search.html" in url:
            ma_so_thue = str(ma_so_thue).strip().replace("'", "")
            ma_tra_cuu = str(ma_tra_cuu).strip()
            mst_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='MST bên bán']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", mst_input)
            mst_input.clear()
            mst_input.send_keys(ma_so_thue)
            
            mtc_input = driver.find_element(By.XPATH, "//input[@placeholder='Mã tra cứu hóa đơn']")
            driver.execute_script("arguments[0].scrollIntoView(true);", mtc_input)
            mtc_input.clear()
            mtc_input.send_keys(ma_tra_cuu)
            
            print("Đang nhấn nút tra cứu...")
            btn = driver.find_element(By.XPATH,"//button[contains(@class, 'webix_button') and contains(text(), 'Tra cứu')]")
            time.sleep(0.5)
            btn.click()
        
        # Tra cứu trên MeInvoice (Misa)
        elif "https://www.meinvoice.vn/tra-cuu/" in url:
            mtc_input = wait.until(EC.presence_of_element_located((By.NAME, "txtCode")))
            driver.execute_script("""
                const header = document.querySelector('.top-header');
                if (header) header.style.display = 'none';
                arguments[0].scrollIntoView({block: 'center'});
            """, mtc_input)
            mtc_input.clear()
            mtc_input.send_keys(ma_tra_cuu)
            
            print("Đang nhấn nút tra cứu...")
            btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSearchInvoice")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(0.5)
            btn.click()

        # Tra cứu trên E-Hóa đơn
        elif "https://van.ehoadon.vn/TCHD?MTC=" in url:
            mtc_input = wait.until(EC.presence_of_element_located((By.ID, "txtInvoiceCode")))
            driver.execute_script("arguments[0].scrollIntoView(true);", mtc_input)
            mtc_input.clear()
            mtc_input.send_keys(ma_tra_cuu)
            
            print("Đang nhấn nút tra cứu...")
            btn = driver.find_element(By.CLASS_NAME, "btnSearch")
            btn.click()
        
        print("Đang chờ kết quả hiển thị...")
        wait.until(EC.presence_of_element_located((By.XPATH, "//body")))
    
    except TimeoutException:
        print("Tra cứu thất bại hoặc trang web phản hồi chậm.")

# Tải file XML hóa đơn về máy
def tai_file_xml(driver, wait, thu_muc_tai_hoa_don, url, ma_tra_cuu):
    try:
        # Tải trên FPT
        if "https://tracuuhoadon.fpt.com.vn/search.html" in url:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[contains(@class, 'mdi-xml')] and contains(text(), 'Tải XML')]")))
            btn.click()
            print("Đã nhấn nút tải XML từ FPT")
            time.sleep(2)

        # Tải trên MeInvoice
        elif "https://www.meinvoice.vn/tra-cuu/" in url:
            btn_menu = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "download")))
            driver.execute_script("arguments[0].scrollIntoView(true);", btn_menu)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", btn_menu)
            
            btn_xml = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "txt-download-xml")))
            driver.execute_script("arguments[0].click();", btn_xml)
            print("Đã chọn tải XML từ MeInvoice")
            time.sleep(2)

        # Tải trên E-Hóa đơn
        elif "https://van.ehoadon.vn/TCHD?MTC=" in url:
            print("Chuyển vào iframe...")
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameViewInvoice")))
            print("Đã vào iframe")
            
            btn_menu = wait.until(EC.presence_of_element_located((By.ID, "btnDownload")))
            ActionChains(driver).move_to_element(btn_menu).perform()
            driver.execute_script("document.querySelector('#divDownloads .dropdown-menu').style.display='block';")
            time.sleep(1)
            
            btn_xml = wait.until(EC.element_to_be_clickable((By.ID, "LinkDownXML")))
            btn_xml.click()
            print("Đã chọn tải XML từ EHoaDon")
            time.sleep(2)
            driver.switch_to.default_content()

    except TimeoutException:
        print("Không tìm thấy nút tải hoặc mã tra cứu sai.")
        return None
    except Exception as e:
        print(f"Lỗi khi tải file: {e}")
        return None

    # Di chuyển file XML vào thư mục riêng theo tên miền
    folder = urlparse(url).netloc.replace("www.", "")
    domain_folder = os.path.join(thu_muc_tai_hoa_don, folder)
    os.makedirs(domain_folder, exist_ok=True)
    
    for _ in range(10):
        files = os.listdir(thu_muc_tai_hoa_don)
        for file in files:
            if file.endswith(".xml"):
                src = os.path.join(thu_muc_tai_hoa_don, file)
                dest = os.path.join(domain_folder, f"{ma_tra_cuu}.xml")
                shutil.move(src, dest)
                print(f"Đã lưu file XML tại: {dest}")
                return dest
        time.sleep(1)
    
    print("Không tìm thấy file XML vừa tải.")
    return None

# Đọc dữ liệu cần thiết từ file XML hóa đơn
def read_invoice_xml(xml_file_path):
    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        hdon_node = root.find(".//HDon")
        invoice_node = hdon_node.find("DLHDon") if hdon_node is not None else None
        
        if invoice_node is None:
            for tag in [".//DLHDon", ".//TDiep", ".//Invoice"]:
                node = root.find(tag)
                if node is not None:
                    invoice_node = node
                    break
            else:
                print(f"Không xác định được node dữ liệu chính trong {os.path.basename(xml_file_path)}")
                return None
        
        def find(path):
            current = invoice_node
            for part in path.split("/"):
                if current is not None:
                    current = current.find(part)
                else:
                    return None
            return current.text if current is not None else None

        # Tìm số tài khoản bán nếu có trong TTKhac
        stk_ban = find("NDHDon/NBan/STKNHang")
        if not stk_ban:
            for info in invoice_node.findall(".//NBan/TTKhac/TTin"):
                if info.findtext("TTruong") == "SellerBankAccount":
                    stk_ban = info.findtext("DLieu")
                    break

        return {
            'Số hóa đơn': find("TTChung/SHDon"),
            'Đơn vị bán hàng': find("NDHDon/NBan/Ten"),
            'Mã số thuế bán': find("NDHDon/NBan/MST"),
            'Địa chỉ bán': find("NDHDon/NBan/DChi"),
            'Số tài khoản bán': stk_ban,
            'Tên người mua hàng': find("NDHDon/NMua/Ten"),
            'Địa chỉ mua': find("NDHDon/NMua/DChi"),
            'MST mua hàng': find("NDHDon/NMua/MST"),
        }
    except Exception as e:
        print(f"Lỗi đọc XML {os.path.basename(xml_file_path)}: {e}")
        return None

# Ghi kết quả vào file Excel
def append_to_excel(filepath, row_data):
    if not os.path.isfile(filepath):
        wb = Workbook()
        ws = wb.active
        ws.title = "Invoices"
        ws.append([
            "STT", "Mã số thuế", "Mã tra cứu", "URL",
            "Số hóa đơn", "Đơn vị bán hàng", "Mã số thuế bán", "Địa chỉ bán", "Số tài khoản bán",
            "Tên người mua hàng", "Địa chỉ mua", "MST mua hàng"
        ])
        wb.save(filepath)
    
    wb = load_workbook(filepath)
    ws = wb.active
    ws.append(row_data)
    wb.save(filepath)

# Chương trình chính
def main():
    input_file = "input.xlsx"
    output_file = "ket_qua_hoa_don.xlsx"
    thu_muc_tai_hoa_don = os.path.join(os.getcwd(), "InvoiceData")

    driver, wait = open_browser(thu_muc_tai_hoa_don)
    df_invoice = pd.read_excel(input_file, dtype=str)

    for index, row in df_invoice.iterrows():
        stt = index + 1
        ma_so_thue = str(row.get("Mã số thuế", "") or "").strip()
        ma_tra_cuu = str(row.get("Mã tra cứu", "") or "").strip()
        url = str(row.get("URL", "") or "").strip()

        if not url or not ma_tra_cuu:
            continue

        print(f"Đang tra cứu: {ma_tra_cuu} | Trang: {url}")
        tra_cuu_hoa_don(driver, wait, ma_so_thue, ma_tra_cuu, url)
        
        xml_path = tai_file_xml(driver, wait, thu_muc_tai_hoa_don, url, ma_tra_cuu)
        
        if xml_path:
            parsed = read_invoice_xml(xml_path)
            if parsed:
                row_data = [stt, ma_so_thue, ma_tra_cuu, url] + list(parsed.values()) + [""]
            else:
                row_data = [stt, ma_so_thue, ma_tra_cuu, url] + [""] * 9 + [os.path.basename(xml_path)]
        else:
            row_data = [stt, ma_so_thue, ma_tra_cuu, url] + [""] * 9 + [""]

        append_to_excel(output_file, row_data)

    driver.quit()
    print(f"Hoàn tất. Kết quả lưu tại: {output_file}")

# Chạy chương trình
if __name__ == "__main__":
    main()
