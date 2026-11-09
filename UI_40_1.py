from datetime import datetime
import io
import re
from difflib import SequenceMatcher
import cv2
import numpy as np
import openpyxl
import pandas as pd
from PIL import Image
import pytesseract
import streamlit as st

import platform
import pytesseract

# กำหนด Path ของ Tesseract-OCR ตามระบบปฏิบัติการ
if platform.system() == "Windows":
  pytesseract.pytesseract.tesseract_cmd = (
      r"C:\Program Files\Tesseract-OCR\tesseract.exe"
  )
else:
  # บน Streamlit Cloud (Linux) ไม่ต้องระบุ path เพราะจะเรียกผ่าน system package ที่ติดตั้งจาก packages.txt
  pass

# ตั้งค่าหน้าตาของ Web Page
st.set_page_config(
    page_title="ระบบตรวจสอบ Certificate จาก Excel - Calibration Lab",
    page_icon="⚖️",
    layout="wide",
)

# --- ปรับแต่ง CSS ---
st.markdown(
    """
    <style>
        .upload-header {
            font-size: 22px !important;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        [data-testid='stFileUploader'] section {
            padding: 40px !important;
            border: 2px dashed #4CAF50 !important;
            background-color: #f9fbf9 !important;
        }
        [data-testid='stFileUploader'] section div {
            font-size: 18px !important;
        }
div.stButton > button {
            background-color: #2e7d32 !important;  /* สีเขียว */
            color: white !important;
            border-radius: 8px !important;
            width: 100% !important;
            padding: 24px 32px !important;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        /* บังคับเจาะจงขยายขนาดตัวอักษรและข้อความภายในปุ่ม Streamlit ให้ใหญ่ขึ้นจริง */
        div.stButton > button p, 
        div.stButton > button span {
            font-size: 26px !important;
            font-weight: bold !important;
            color: white !important;
        }

        div.stButton > button:hover {
            background-color: #1b5e20 !important;  /* สีเขียวเข้มขึ้นเมื่อเอาเมาส์ชี้ */
        }
        }
        div.stButton > button:hover {
            background-color: #ff2b2b !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# --- ฟังก์ชันทำความสะอาดข้อความเพื่อเปรียบเทียบ ---
def clean_str(text):
  if not text or pd.isna(text):
    return ""
  return re.sub(r"\s+", "", str(text)).strip().upper()


def clean_address(text):
  if not text or pd.isna(text):
    return ""
  text = re.sub(r"e-?mail\s*:.*", "", str(text), flags=re.IGNORECASE)
  text = re.sub(r"\bThailand\b", "", str(text), flags=re.IGNORECASE)
  return re.sub(r"[\s,]", "", text).upper()


# --- แสดงตำแหน่งข้อความที่ไม่ตรงกัน (ไม่เปลี่ยนเกณฑ์ PASS / FAIL เดิม) ---
def show_text_difference(container, excel_value, ocr_value, address=False):
  normalize = clean_address if address else clean_str
  a = normalize(excel_value)
  b = normalize(ocr_value)
  if not b:
    container.caption("รายละเอียด: OCR ไม่พบข้อความสำหรับเปรียบเทียบ")
    return
  if not a:
    container.caption("รายละเอียด: Excel ไม่มีข้อความสำหรับเปรียบเทียบ")
    return
  if a == b:
    return

  # เปรียบเทียบอักขระหลัง normalization เดียวกับที่ใช้ตัดสิน PASS/FAIL
  differences = []
  for tag, i1, i2, j1, j2 in SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
    if tag == "equal":
      continue
    left = a[i1:i2] or "(ไม่มี)"
    right = b[j1:j2] or "(ไม่มี)"
    differences.append(f"Excel: `{left}` | OCR: `{right}` (ตำแหน่ง Excel {i1 + 1}–{i2})")
  container.caption("ส่วนที่ไม่ตรงกัน (หลังตัดช่องว่าง/ปรับตัวพิมพ์):")
  for diff in differences[:12]:
    container.caption("• " + diff)
  if len(differences) > 12:
    container.caption(f"…และอีก {len(differences) - 12} จุด")


def show_cert_difference(container, value):
  expected = "MP-0XXX-XX (X เป็นตัวเลข)"
  issues = []
  if not re.fullmatch(r"MP-", str(value)[:3]):
    issues.append("Prefix ต้องเป็น MP-")
  if len(str(value).strip()) != 10:
    issues.append(f"ความยาว {len(str(value).strip())} ตัวอักษร (ต้องเป็น 10)")
  if not re.fullmatch(r"MP-0\d{3}-\d{2}", str(value).strip()):
    issues.append(f"ค่าปัจจุบัน: `{value}` | รูปแบบที่กำหนด: `{expected}`")
  for issue in issues:
    container.caption("• " + issue)


def validate_cert_no(cert_str):
  pattern = r"^MP-0\d{3}-\d{2}$"
  if not cert_str or pd.isna(cert_str):
    return False
  return bool(re.match(pattern, str(cert_str).strip()))


# --- ฟังก์ชันสกัดข้อความจากภาพ Work Order ด้วย OCR (ปรับปรุงใหม่ให้แม่นยำและยืดหยุ่นสูงขึ้น) ---
def extract_work_order_data(pil_img):
  try:
    open_cv_image = np.array(pil_img)
    gray = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)

    ocr_text = pytesseract.image_to_string(thresh, lang="eng")
    lines = [line.strip() for line in ocr_text.split("\n") if line.strip()]
    full_text = " ".join(lines)

    ref_data = {
        "cal_id": "",
        "service_item": "",
        "manufacturer": "",
        "customer": "",
        "address": "",
        "model": "",
        "serial_no": "",
    }

    # 1. Cal ID (รองรับทั้ง CAL_ID, CAL ID และกวาดหารูปแบบรหัส MEC... ทั่วทั้งภาพ)
    cal_match = re.search(
        r"CAL[_\s]*ID\s*[:\s]*([A-Za-z0-9\-\/]+)", full_text, re.IGNORECASE
    )
    if cal_match:
      ref_data["cal_id"] = cal_match.group(1).strip()
    else:
      # Fallback: ค้นหารูปแบบรหัสมาตรฐาน เช่น MEC1105-01/26 จากข้อความทั้งหมดโดยตรง
      id_cand = re.search(r"(MEC\d{4}-\d{2}/\d{2})", full_text, re.IGNORECASE)
      if id_cand:
        ref_data["cal_id"] = id_cand.group(1).strip()
      else:
        # Fallback สำรองแบบกวาดหาคำว่า MEC หรือ CAL ในแต่ละบรรทัด
        for line in lines:
          if "MEC" in line.upper() or "CAL" in line.upper():
            parts = re.split(r"[:\s]+", line)
            for p in parts:
              if "MEC" in p.upper() or "/" in p:
                ref_data["cal_id"] = p.strip()
                break
            if ref_data["cal_id"]:
              break

    # 2. Service Item Description / Equipment
    eq_match = re.search(
        r"Equipment\s*/\s*Sample\s*No\.\s*[:\s]*(.*?)(?=Manufacturer|$)",
        full_text,
        re.IGNORECASE,
    )
    if eq_match:
      raw_service = eq_match.group(1).strip()
      ref_data["service_item"] = raw_service.split("(")[0].strip()

    # 3. Customer Name (ปรับปรุงขอบเขตให้หยุดทันทีเมื่อเจอคำว่า Address หรือเครื่องหมายขึ้นบรรทัดใหม่/คีย์เวิร์ดถัดไป)
    cust_match = re.search(
        r"Customer(?:\s*Name)?\s*[:\s]*(.*?)(?=\s+Address\s*:|\s+On-site|\s+Measurement|$)",
        full_text,
        re.IGNORECASE,
    )
    if cust_match:
      raw_cust = cust_match.group(1).strip()
      # ตัดข้อความส่วนเกินหากมีคำว่า On-site หรืออื่นๆ ติดมา
      for stop_word in [
          "On-site",
          "Measurement",
          "RQ_ID",
          "CAL_ID",
          "Address",
      ]:
        if stop_word in raw_cust:
          raw_cust = raw_cust.split(stop_word)[0].strip()
      ref_data["customer"] = raw_cust

    # 4. Customer Address - หยุดที่ label ของข้อมูลถัดไป ไม่ใช่คำว่า Tel ใน Tellhow
    addr_match = re.search(
        r"\bAddress\s*:\s*(.*?)\s*(?=\be[\s-]*mail\s*:|\bTel\s*\.?\s*:|\bMobile\s*:|\bEquipment\s*/\s*Sample\s*No\.?\s*:|$)",
        full_text,
        re.IGNORECASE,
    )
    if addr_match:
      ref_data["address"] = addr_match.group(1).strip()

    # 5. Manufacturer
    manuf_match = re.search(
        r"Manufacturer\s*[:\s]*([A-Za-z0-9\s]+?)(?=Model|$)",
        full_text,
        re.IGNORECASE,
    )
    if manuf_match:
      ref_data["manufacturer"] = manuf_match.group(1).strip()

    # 6. Model - รองรับชื่อรุ่นที่มีช่องว่าง เช่น FLUKE 922
    model_match = re.search(
        r"\bModel\s*:\s*(.*?)\s*(?=\bSerial\s*No\.?|\bRange\s*:|$)",
        full_text,
        re.IGNORECASE,
    )
    if model_match:
      ref_data["model"] = model_match.group(1).strip()

    # 7. Serial Number (ปรับปรุงใหม่ให้จับเฉพาะบรรทัด Serial No. โดยไม่สับสนกับ Sample No. ของ Equipment)
    sn_match = re.search(
        r"Serial\s*No[.\s]*(?:/\s*Sample\s*No\.)?\s*[:\s]*([A-Za-z0-9\-]+)",
        full_text,
        re.IGNORECASE,
    )
    if sn_match:
      ref_data["serial_no"] = sn_match.group(1).strip()
    else:
      # วิธีสำรอง: ค้นหาโดยไล่เช็คทีละบรรทัดที่ขึ้นต้นด้วยคำว่า Serial
      for line in lines:
        if "SERIAL" in line.upper():
          parts = re.split(r"[:]", line)
          if len(parts) > 1:
            val = parts[1].strip()
            if val and val.lower() != "sample":
              ref_data["serial_no"] = val
              break
    
    return ref_data, ocr_text
  except Exception as e:
    return {}, str(e)


# --- UI Header ---
st.title("⚖️ (ระบบเก่า) Certificate Verification - Pressure & Vacuum Lab 🏷️")
st.caption(
    "ระบบตรวจสอบความถูกต้องข้อมูลในตาราง Excel และรูปภาพ Work Order ด้วย OCR"
)
st.divider()

st.markdown(
    '<p class="upload-header">📁 กรุณาอัปโหลดไฟล์ Excel (เช่น .xlsx)</p>',
    unsafe_allow_html=True,
)
uploaded_excel = st.file_uploader(
    "", type=["xlsx"], key="excel_file", label_visibility="collapsed"
)

st.divider()

if uploaded_excel:
  if st.button(
      "🚀 เริ่มทำการตรวจสอบเอกสาร (Start Verification)", use_container_width=True
  ):
    try:
      # 1. อ่าน Workbook
      wb = openpyxl.load_workbook(uploaded_excel, data_only=True)

      # 2. อ่านข้อมูลตาราง Excel Sheet Page1
      df_page1 = pd.read_excel(uploaded_excel, sheet_name="Page1", header=None)

      cert_no_val = str(df_page1.iloc[2, 4]).strip()
      cal_id_val = str(df_page1.iloc[4, 4]).strip()
      service_item_val = str(df_page1.iloc[5, 2]).strip()
      manufacturer_val = str(df_page1.iloc[7, 2]).strip()
      model_excel_val = str(df_page1.iloc[8, 2]).strip()
      sn_excel_val = str(df_page1.iloc[9, 2]).strip()
      cust_excel_val = str(df_page1.iloc[10, 2]).strip()
      addr_excel_val = (
          str(df_page1.iloc[11, 2]) + " " + str(df_page1.iloc[12, 2])
      ).strip()
      meas_date_val = df_page1.iloc[14, 2]

      # 3. อ่าน Report Date จากชีต Certificate ที่มองเห็นได้เท่านั้น
      # จับคู่ Cal ID ในแถวเดียวกัน และเลือกวันที่ทางขวาของ Cal ID
      target_cert_sheet = None
      report_date_val = "ไม่พบข้อมูล Report Date"
      normalized_cal_id = clean_str(cal_id_val)

      preferred_sheets = ["Cert TISI-P1", "Cert CIPM-P1"]
      visible_cert_sheets = [
          name for name in wb.sheetnames
          if wb[name].sheet_state == "visible"
          and name.strip() in preferred_sheets
      ]
      visible_cert_sheets.sort(
          key=lambda name: preferred_sheets.index(name.strip())
      )

      for sheet_name in visible_cert_sheets:
        sheet = wb[sheet_name]
        found_date = None
        for row in sheet.iter_rows():
          for cell in row:
            if clean_str(cell.value) != normalized_cal_id or not normalized_cal_id:
              continue
            # วันที่ต้องอยู่ทางขวาของ Cal ID ในแถวเดียวกัน
            for col_idx in range(cell.column + 1, sheet.max_column + 1):
              candidate = sheet.cell(row=cell.row, column=col_idx).value
              if isinstance(candidate, (datetime, pd.Timestamp)):
                found_date = candidate
                break
              if isinstance(candidate, str) and candidate.strip():
                text_date = candidate.strip()
                # รับเฉพาะข้อความที่มีลักษณะเป็นวันที่ ไม่ตีความเลขอื่นเป็นวันที่
                if re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}", text_date):
                  try:
                    found_date = pd.to_datetime(text_date, dayfirst=True)
                    break
                  except (ValueError, TypeError):
                    pass
                elif re.search(r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\b", text_date, re.IGNORECASE):
                  try:
                    found_date = pd.to_datetime(text_date)
                    break
                  except (ValueError, TypeError):
                    pass
            if found_date is not None:
              break
          if found_date is not None:
            break
        if found_date is not None:
          target_cert_sheet = sheet_name
          report_date_val = found_date
          break

      col1, col2, c3 = st.columns([2, 2, 1])

      # 4. อ่านภาพ Work Order และรัน OCR
      sheet_page1 = wb["Page1"]
      pil_img = None
      ocr_dict = {}

      if sheet_page1._images:
        img = sheet_page1._images[0]
        img_data = img._data() if callable(img._data) else img._data
        pil_img = Image.open(io.BytesIO(img_data))
        ocr_dict, raw_ocr_text = extract_work_order_data(pil_img)

      ref_cal_id = ocr_dict.get("cal_id", "")
      ref_service = ocr_dict.get("service_item", "")
      ref_manuf = ocr_dict.get("manufacturer", "")
      ref_cust = ocr_dict.get("customer", "")
      ref_addr = ocr_dict.get("address", "")
      ref_model = ocr_dict.get("model", "")
      ref_sn = ocr_dict.get("serial_no", "")

      # --- Dashboard แสดงผล ---
      st.header("📋 สรุปผลการตรวจสอบ (Verification Results)")
      col_view1, col_view2 = st.columns(2)

      with col_view1:
        st.subheader("📊 ข้อมูลจากตาราง Excel (Certificate Data)")
        st.write(f"**Certificate No.:** {cert_no_val}")
        st.write(f"**Cal ID:** {cal_id_val}")
        st.write(f"**Service Item Description:** {service_item_val}")
        st.write(f"**Manufacturer:** {manufacturer_val}")
        st.write(f"**Customer Name:** {cust_excel_val}")
        st.write(f"**Customer Address:** {addr_excel_val}")
        st.write(f"**Model:** {model_excel_val}")
        st.write(f"**Serial Number:** {sn_excel_val}")
        st.write(f"**Measurement Date:** {meas_date_val}")
        st.write(
            f"**Report Date ({target_cert_sheet if target_cert_sheet else 'N/A'}):**"
            f" {report_date_val}"
        )

      with col_view2:
        st.subheader("🖼️ รูปภาพ Work Order & OCR Extracted Ref")
        if pil_img:
          st.image(
              pil_img,
              caption="Work Order Image (สกัดด้วย Tesseract OCR)",
              use_container_width=True,
          )
        else:
          st.warning("ไม่พบรูปภาพ Work Order ใน Sheet Page1")

      st.divider()
      st.subheader(
          "✅ ผลการตรวจสอบความถูกต้อง (Check List & Consistency Verification)"
      )

      # [1] Certificate No. Check
      st.markdown("### [1] Certificate No. Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel Certificate No.:** {cert_no_val}")
      c2.write(f"**Required Format:** MP-0XXX-XX (10 ตัวอักษร)")
      if validate_cert_no(cert_no_val):
        c3.success("PASS ✅")
      else:
        c3.error("FAIL ❌ (รูปแบบ Certificate No. ไม่ถูกต้อง)")
        show_cert_difference(c3, cert_no_val)

      # [2] Cal ID Check
      st.markdown("### [2] Cal ID Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel Cal ID:** {cal_id_val}")
      c2.write(f"**File Ref Cal ID (OCR):** {ref_cal_id}")
      if ref_cal_id and clean_str(cal_id_val) == clean_str(ref_cal_id):
        c3.success("PASS ✅")
      else:
        c3.error("FAIL ❌ (Cal ID ไม่ตรงกันหรือพบข้อผิดพลาด)")
        show_text_difference(c3, cal_id_val, ref_cal_id)

      # [3] Service Item Description Check
      st.markdown("### [3] Service Item Description Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel Description:** {service_item_val}")
      c2.write(f"**Work Order Ref (OCR):** {ref_service}")
      if ref_service and clean_str(service_item_val) == clean_str(ref_service):
        c3.success("PASS ✅")
      else:
        c3.error(
            "FAIL ❌ (รายละเอียดเครื่องมือไม่ตรงกัน หรือมีข้อความส่วนเกิน)"
        )
        show_text_difference(c3, service_item_val, ref_service)

      # [4] Manufacturer Check
      st.markdown("### [4] Manufacturer Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel Manufacturer:** {manufacturer_val}")
      c2.write(f"**Work Order Ref (OCR):** {ref_manuf}")
      if ref_manuf and clean_str(manufacturer_val) == clean_str(ref_manuf):
        c3.success("PASS ✅")
      else:
        c3.error("FAIL ❌ (ผู้ผลิตไม่ตรงกับ Work Order)")
        show_text_difference(c3, manufacturer_val, ref_manuf)

      # [5] Customer Name Check
      st.markdown("### [5] Customer Name Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel:** {cust_excel_val}")
      c2.write(f"**File Ref (OCR):** {ref_cust}")
      if ref_cust and clean_str(cust_excel_val) == clean_str(ref_cust):
        c3.success("PASS ✅")
      else:
        c3.error("FAIL ❌ (ชื่อลูกค้าไม่ตรงกันหรือไม่พบข้อมูลจาก OCR)")
        show_text_difference(c3, cust_excel_val, ref_cust)

      # [6] Customer Address Check
      st.markdown("### [6] Customer Address Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel Address:** {addr_excel_val}")
      c2.write(f"**File Ref Address (OCR):** {ref_addr}")
      if ref_addr and clean_address(addr_excel_val) == clean_address(ref_addr):
        c3.success("PASS ✅")
      else:
        c3.error(
            "FAIL ❌ (ที่อยู่ระหว่าง Excel และรูป Work Order ไม่ตรงกัน)"
        )
        show_text_difference(c3, addr_excel_val, ref_addr, address=True)

      # [7] Model Check
      st.markdown("### [7] Model / Type Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel Model:** {model_excel_val}")
      c2.write(f"**File Ref Model (OCR):** {ref_model}")
      if ref_model and clean_str(model_excel_val) == clean_str(ref_model):
        c3.success("PASS ✅")
      else:
        c3.error("FAIL ❌ (รุ่นไม่ตรงกัน)")
        show_text_difference(c3, model_excel_val, ref_model)

      # [8] Serial Number Check
      st.markdown("### [8] Serial Number Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Excel S/N:** {sn_excel_val}")
      c2.write(f"**File Ref S/N (OCR):** {ref_sn}")
      if ref_sn and clean_str(sn_excel_val) == clean_str(ref_sn):
        c3.success("PASS ✅")
      else:
        c3.error("FAIL ❌ (S/N ไม่ตรงกัน)")
        show_text_difference(c3, sn_excel_val, ref_sn)

     # [9] Report Date vs Measurement Date Sequence Check
      st.markdown("### [9] Report Date & Measurement Date Sequence Check")
      c1, c2, c3 = st.columns([2, 2, 1])
      c1.write(f"**Measurement Date:** {meas_date_val}")
      c2.write(f"**Report Date (Found Near Cal ID):** {report_date_val}")

      if (
          not target_cert_sheet
          or report_date_val == "ไม่พบข้อมูล Report Date"
      ):
        c3.info("SKIPPED ℹ️ (ไม่พบข้อมูลวันที่ในชีทรายงาน)")
      else:
        try:
          # ทำความสะอาดรูปแบบวันที่ (เช่น เพิ่มช่องว่างหลังคอมมาหากมี)
          str_meas = str(meas_date_val).replace(",", ", ")
          str_report = str(report_date_val)

          dt_meas = pd.to_datetime(str_meas)
          dt_report = pd.to_datetime(str_report)

          # เงื่อนไขที่ถูกต้อง: วันวัดค่า (dt_meas) ต้องมาก่อนหรือเกิดขึ้นในวันเดียวกับวันออกรายงาน (dt_report)
          if dt_meas <= dt_report:
            c3.success("PASS ✅")
          else:
            c3.error(
                "FAIL ❌ (Measurement Date เกิดขึ้นทีหลังวัน Report Date)"
            )
            c3.caption(f"Measurement Date: {dt_meas.date()} ช้ากว่า Report Date: {dt_report.date()} อยู่ {(dt_meas.normalize() - dt_report.normalize()).days} วัน")
        except Exception as date_err:
          c3.error(f"FAIL ❌ (รูปแบบวันที่ไม่ถูกต้อง: {date_err})")

    except Exception as e:
      st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")

else:
  st.info(
      "💡 กรุณาอัปโหลดไฟล์ Excel ที่มีข้อมูลใน Sheet 'Page1' และรูปภาพ Work Order"
  )