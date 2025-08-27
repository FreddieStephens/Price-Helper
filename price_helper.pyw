import os
import subprocess
import logging
import psutil
import win32com.client
import pythoncom
import urllib.parse
from flask import Flask, request, jsonify
from shutil import copyfile

# -----------------------
# Logging setup (store logs OUTSIDE SharePoint so sync never blocks)
# -----------------------
log_path = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "sheet_helper.log"
)
logging.basicConfig(
    filename=log_path,
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s"
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# -----------------------
# Config
# -----------------------
BASE_DIR = os.path.join(
    os.path.expanduser("~"),
    r"Synergy Lifting\Synergy Lifting Hub - Documents\1. Project Paperwork\a. Quotes & Pricing"
)
PORT = 54007

# -----------------------
# Helpers
# -----------------------
def is_file_open(file_name):
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        try:
            if proc.info['name'] and file_name.lower() in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def open_file(path):
    try:
        if not os.path.exists(path):
            logger.warning(f"File not found: {path}")
            return
        os.startfile(path)
        logger.info(f"Opened file: {path}")
    except Exception as e:
        logger.error(f"Failed to open {path}: {e}")

def open_folder(folder):
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        target_name = os.path.basename(os.path.normpath(folder)).lower()
        found = False

        logger.info(f"Looking for existing Explorer window for folder '{target_name}'")

        for window in shell.Windows():
            try:
                if window.LocationURL:
                    url = window.LocationURL.replace("file:///", "")
                    path = urllib.parse.unquote(url).replace("/", "\\")
                    current_name = os.path.basename(os.path.normpath(path)).lower()

                    if current_name == target_name:
                        try:
                            window.Visible = True
                            window.Focus()
                            logger.info(f"Reused existing Explorer window for {folder}")
                        except Exception as e:
                            logger.warning(f"Matched {folder} but Focus() failed: {e}")
                            logger.info(f"Reused existing Explorer window for {folder} (without focus)")
                        found = True
                        break
            except Exception as e:
                logger.warning(f"Error while scanning Explorer windows: {e}")
                continue

        if not found:
            subprocess.Popen(f'explorer "{folder}"')
            logger.info(f"Opened new Explorer window for {folder}")

    except Exception as e:
        logger.error(f"Failed to open folder {folder}: {e}")
    finally:
        pythoncom.CoUninitialize()

def force_onedrive_sync(path):
    """
    Trigger OneDrive to upload the given file immediately.
    """
    try:
        onedrive_exe = os.path.expandvars(r"%localappdata%\Microsoft\OneDrive\onedrive.exe")
        if os.path.exists(onedrive_exe):
            subprocess.Popen([onedrive_exe, "/triggerupload", path])
            logger.info(f"Triggered OneDrive sync for {path}")
        else:
            logger.warning("OneDrive executable not found, could not trigger sync")
    except Exception as e:
        logger.error(f"Failed to trigger OneDrive sync: {e}")

def kill_stray_processes():
    """
    Kill any leftover Excel or Python processes to avoid file locks.
    """
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        try:
            name = proc.info['name']
            if name:
                if "EXCEL.EXE" in name.upper():
                    proc.kill()
                    logger.info(f"Killed stray Excel process (PID {proc.info['pid']})")
        except Exception:
            continue

def update_excel_fields(path, oppid, title, miles, close, owner):
    """
    Use Excel COM automation to reliably write into merged/formatted cells.
    """
    try:
        pythoncom.CoInitialize()
        logger.info("Starting Excel COM update...")

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(path)

        # Salesforce Meta
        try:
            ws_meta = wb.Sheets("Salesforce Meta")
            ws_meta.Range("B2").Value = oppid
            logger.info(f"Wrote oppid '{oppid}' to Salesforce Meta!B2")
        except Exception as e:
            logger.error(f"Failed writing to Salesforce Meta sheet: {e}", exc_info=True)

        # Pricing
        try:
            ws_pricing = wb.Sheets("Pricing")
            ws_pricing.Range("A2").Value = title
            logger.info(f"Wrote title '{title}' to Pricing!A2")

            ws_pricing.Range("F2").Value = miles
            logger.info(f"Wrote miles '{miles}' to Pricing!F2")

            ws_pricing.Range("H3").Value = owner
            logger.info(f"Wrote owner '{owner}' to Pricing!H3")

            ws_pricing.Range("I3").Value = close
            logger.info(f"Wrote close '{close}' to Pricing!I3")
        except Exception as e:
            logger.error(f"Failed writing to Pricing sheet: {e}", exc_info=True)

        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        logger.info(f"Successfully updated and saved Excel file {path}")

        # ✅ Force upload and cleanup
        force_onedrive_sync(path)
        kill_stray_processes()

    except Exception as e:
        logger.error(f"Failed updating Excel via COM: {e}", exc_info=True)
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

# -----------------------
# Routes
# -----------------------
@app.route("/run_price", methods=["POST"])
def run_price():
    try:
        data = request.get_json(force=True)
        proj_id = data.get("proj_id")
        oppid = data.get("oppid")
        title = data.get("title")
        miles = data.get("miles")
        close = data.get("close")
        owner = data.get("owner")

        folder = os.path.join(BASE_DIR, f"QP {proj_id}")
        os.makedirs(folder, exist_ok=True)

        wb_path = os.path.join(folder, f"P {proj_id}.xlsm")
        first_time = False

        if not os.path.exists(wb_path):
            template_path = os.path.join(BASE_DIR, "i. Templates", "Pricebook Template.xlsm")
            if os.path.exists(template_path):
                copyfile(template_path, wb_path)
                first_time = True
                logger.info(f"Copied template to {wb_path}")
            else:
                logger.error(f"Template not found: {template_path}")

        if first_time:
            update_excel_fields(wb_path, oppid, title, miles, close, owner)

        open_folder(folder)
        if os.path.exists(wb_path):
            open_file(wb_path)

        return jsonify({"status": "ok", "file": wb_path})
    except Exception as e:
        logger.error(f"Error in /run_price: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route("/run_quote", methods=["POST"])
def run_quote():
    try:
        data = request.get_json(force=True)
        proj_id = data.get("proj_id")
        folder = os.path.join(BASE_DIR, f"QP {proj_id}")
        os.makedirs(folder, exist_ok=True)

        # Find latest Word doc by modified date
        docs = [f for f in os.listdir(folder) if f.lower().endswith(".docx")]
        doc_path = None

        if docs:
            docs.sort(key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)
            doc_path = os.path.join(folder, docs[0])
            logger.info(f"Latest Word doc found: {doc_path}")
        else:
            logger.info(f"No Word docs found in {folder}")

        open_folder(folder)
        if doc_path:
            open_file(doc_path)
            logger.info(f"Opened latest Word doc: {doc_path}")

        return jsonify({"status": "ok", "file": doc_path})
    except Exception as e:
        logger.error(f"Error in /run_quote: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route("/health")
def health():
    return jsonify({"status": "ok"})

# -----------------------
# Entry
# -----------------------
if __name__ == "__main__":
    logger.info(f"Starting sheet_helper on port {PORT}")
    app.run(port=PORT, debug=False)
