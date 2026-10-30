# ============================================================
# TREMENDOUS BALANCE EXTRACTOR – NODRIVER + STREAMLIT UI
# ============================================================

import asyncio
import os
import re
import tempfile
import traceback
import openpyxl
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import nodriver as uc
import nest_asyncio

# Apply nest_asyncio for Streamlit's async environment
nest_asyncio.apply()

# Ensure DISPLAY is set for the virtual screen (Xvfb)
if 'DISPLAY' not in os.environ:
    os.environ['DISPLAY'] = ':99'

CHROME_PATH = "/usr/bin/google-chrome"
INTER_REQUEST_DELAY_SECONDS = 2

# --- SESSION STATE INITIALIZATION ---
if "targets" not in st.session_state: st.session_state.targets = None          
if "results" not in st.session_state: st.session_state.results = None          
if "uploaded_file" not in st.session_state: st.session_state.uploaded_file = None    
if "processing_complete" not in st.session_state: st.session_state.processing_complete = False
if "file_name" not in st.session_state: st.session_state.file_name = None

# --- HELPER FUNCTIONS ---
def normalize_cell_value(value):
    if value is None: return ""
    if isinstance(value, int): return str(value)
    if isinstance(value, float): return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()

def get_hyperlink_or_url(cell):
    try:
        if cell.hyperlink and cell.hyperlink.target: return str(cell.hyperlink.target).strip()
    except: pass
    value = cell.value
    if value is None: return None
    value_str = str(value).strip()
    if value_str.lower().startswith(("http://", "https://")): return value_str
    return None

def extract_balance(text):
    if not text: return None
    t = text.replace("\u00a0", " ")
    m = re.search(r"Available\s+Balance[\s\S]{0,50}?\$\s*([\d,]+\.\d{2})", t, re.I)
    if m: return f"{float(m.group(1).replace(',', '')):.2f}"
    return None

# --- NODRIVER PROCESSING LOGIC ---
async def process_one_url(driver, url):
    if not url: return None, "No URL"
    try:
        tab = await driver.get(url)
    except Exception:
        raise Exception("DriverDead")

    try:
        await tab.sleep(3)
        try:
            page_title = await tab.evaluate("document.title")
        except: page_title = ""

        if "Just a moment" in page_title or "Performing security verification" in page_title:
            try:
                await tab.verify_cf()
                await tab.sleep(5) 
            except: pass

        balance = None
        for _ in range(15):
            try:
                page_text = await tab.evaluate("document.body ? document.body.innerText : ''")
                balance = extract_balance(page_text)
                if balance: break
            except: pass
            await tab.sleep(1)

        return (balance, "Success") if balance else (None, "Balance Not Found")

    except Exception:
        raise Exception("DriverDead")

async def start_browser():
    return await uc.start(
        browser_executable_path=CHROME_PATH,
        headless=False, 
        browser_args=["--window-size=1920,1080", "--no-sandbox", "--disable-dev-shm-usage"]
    )

async def run_processor(targets, progress_callback):
    results = []
    total = len(targets)
    driver = await start_browser()
    
    try:
        for idx, target in enumerate(targets, 1):
            progress_callback(idx, total, target["reward_id"])
            url = target["url"]
            
            if not url:
                results.append({**target, "balance": "N/A", "status": "No URL"})
                continue

            try:
                balance, status = await process_one_url(driver, url)
            except Exception:
                # Crash Recovery Protocol
                try: driver.stop()
                except: pass
                await asyncio.sleep(2)
                driver = await start_browser()
                try: balance, status = await process_one_url(driver, url)
                except: balance, status = None, "Crash Retry Failed"

            results.append({
                **target,
                "balance": balance if balance is not None else "N/A",
                "status": status
            })
            
            if idx < total: await asyncio.sleep(INTER_REQUEST_DELAY_SECONDS)
    finally:
        driver.stop()
    return results

def export_to_excel(results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Balances"
    headers = ["Reward ID", "Card Number", "Extracted Balance", "Overall Status"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    link_font = Font(color="0563C1", underline="single")
    regular_font = Font()
    
    for col in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{len(results)+1}"
    
    for row_idx, item in enumerate(results, start=2):
        cell_a = ws.cell(row=row_idx, column=1)
        cell_a.value = item["reward_id"]
        if item.get("url"):
            cell_a.hyperlink = item["url"]
            cell_a.font = link_font
        else: cell_a.font = regular_font
            
        cell_b = ws.cell(row=row_idx, column=2)
        cell_b.value = item["card_number"]
        if item.get("card_url"):
            cell_b.hyperlink = item["card_url"]
            cell_b.font = link_font
        else: cell_b.font = regular_font
            
        ws.cell(row=row_idx, column=3).value = item["balance"]
        ws.cell(row=row_idx, column=4).value = item["status"]
        
    for col in ws.columns:
        max_len = 0
        for cell in col:
            try: max_len = max(max_len, len(str(cell.value or "")))
            except: pass
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len+4, 15), 50)
        
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        return tmp.name

# --- STREAMLIT UI ---
st.set_page_config(page_title="Tremendous Balance Extractor", layout="wide")
st.title("💳 Tremendous Balance Extractor")
st.markdown("Upload your Excel file (Columns: **A = Reward ID**, **B = Card Number**) and click **Process**.")

col_clear, _ = st.columns([1, 5])
with col_clear:
    if st.button("🗑️ Clear All Data", use_container_width=True):
        for key in ["targets", "results", "uploaded_file", "processing_complete", "file_name"]:
            if key in st.session_state: del st.session_state[key]
        st.rerun()

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"], key="file_uploader")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    if st.session_state.uploaded_file != file_bytes:
        st.session_state.uploaded_file = file_bytes
        st.session_state.file_name = uploaded_file.name
        st.session_state.targets = None
        st.session_state.results = None
        st.session_state.processing_complete = False

    if st.session_state.targets is None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
            tmp_input.write(st.session_state.uploaded_file)
            input_path = tmp_input.name

        wb_input = openpyxl.load_workbook(input_path, data_only=False)
        sheet = wb_input["Main List"] if "Main List" in wb_input.sheetnames else wb_input.active
        targets = []
        for row_num, row in enumerate(sheet.iter_rows(min_row=2, max_col=2), start=2):
            cell_a, cell_b = row[0], row[1]
            if cell_a.value is None and cell_b.value is None: continue
            reward_id = normalize_cell_value(cell_a.value) or "N/A"
            card_number = normalize_cell_value(cell_b.value) or "N/A"
            reward_url = get_hyperlink_or_url(cell_a)
            card_url = get_hyperlink_or_url(cell_b)
            if not reward_url and card_url: reward_url = card_url
            targets.append({
                "row": row_num, "reward_id": reward_id, "card_number": card_number,
                "url": reward_url, "card_url": card_url,
            })
        wb_input.close()
        os.unlink(input_path)
        st.session_state.targets = targets
        st.session_state.processing_complete = False

    st.success(f"✅ Loaded {len(st.session_state.targets)} records from {st.session_state.file_name}")

    if st.session_state.results is not None and st.session_state.processing_complete:
        results = st.session_state.results
        st.subheader("📊 Results Summary")
        success = sum(1 for r in results if r["status"] == "Success")
        not_found = sum(1 for r in results if r["status"] == "Balance Not Found")
        errors = sum(1 for r in results if "Error" in r["status"] or "Crash" in r["status"])

        col1, col2, col3 = st.columns(3)
        col1.metric("✅ Success", success)
        col2.metric("⚠️ Balance Not Found", not_found)
        col3.metric("❌ Errors / Crash", errors)

        df = pd.DataFrame(results)
        df["Reward ID"] = df.apply(lambda row: f'<a href="{row["url"]}" target="_blank">{row["reward_id"]}</a>' if row["url"] else row["reward_id"], axis=1)
        df["Card Number"] = df.apply(lambda row: f'<a href="{row["card_url"]}" target="_blank">{row["card_number"]}</a>' if row["card_url"] else row["card_number"], axis=1)
        df_display = df[["Reward ID", "Card Number", "balance", "status"]]
        df_display.columns = ["Reward ID", "Card Number", "Balance", "Status"]
        
        st.write("### Results Table")
        st.markdown(df_display.to_html(escape=False, index=False), unsafe_allow_html=True)

        excel_path = export_to_excel(results)
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 Download Report (Excel)",
                data=f,
                file_name="Tremendous_Balances_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        os.unlink(excel_path)

        if st.button("🔄 Process Again (re-run all URLs)", use_container_width=True):
            st.session_state.results = None
            st.session_state.processing_complete = False
            st.rerun()

    if not st.session_state.processing_complete:
        if st.button("🚀 Process All URLs", type="primary", use_container_width=True):
            targets = st.session_state.targets
            progress_bar = st.progress(0)
            status_text = st.empty()

            def progress_callback(current, total, reward_id):
                progress_bar.progress(current / total)
                status_text.text(f"Processing {current}/{total}: {reward_id[:30]}...")

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                results = loop.run_until_complete(run_processor(targets, progress_callback))
            except Exception as e:
                st.error(f"❌ An error occurred: {e}\n\n{traceback.format_exc()}")
            finally:
                loop.close()

            st.session_state.results = results
            st.session_state.processing_complete = True
            progress_bar.empty()
            status_text.empty()
            st.rerun()
else:
    st.info("Please upload an Excel file to begin.")
