import win32com.client
import os
from datetime import datetime, timedelta
import pythoncom

def get_cutoff_message(check_days):
    threshold_date = datetime.now() - timedelta(days=int(check_days))
    return f"Only links modified on or after {threshold_date.strftime('%Y-%m-%d %H:%M:%S')} will be updated."

def run_excel_link_update(options: dict, print_func=None):
    CHECK_DAYS = int(options.get("CHECK_DAYS", 14))
    LOG_DIR = options.get("LOG_DIR", r"D:\Pzone\Log")
    SHOW_FULL_PATH = bool(options.get("SHOW_FULL_PATH", False))
    SHOW_LINK = bool(options.get("SHOW_LINK", False))
    SHOW_LAST_MODIFIED = bool(options.get("SHOW_LAST_MODIFIED", False))
    SHOW_STATUS = bool(options.get("SHOW_STATUS", False))
    SAVE_LOG = bool(options.get("SAVE_LOG", True))
    if SAVE_LOG:
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR, exist_ok=True)
        LOG_FILE = os.path.join(LOG_DIR, f"excel_link_update_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        log_file_handler = open(LOG_FILE, "a", encoding="utf-8")
    else:
        log_file_handler = None

    def print_log(msg):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{ts} | {msg}"
        if print_func:
            print_func(msg)
        else:
            print(line)
        if SAVE_LOG and log_file_handler:
            log_file_handler.write(line + "\n")
            log_file_handler.flush()

    def setup_excel():
        try:
            pythoncom.CoInitialize()
            return win32com.client.Dispatch("Excel.Application")
        except Exception as e:
            print_log(f"Excel initialization failed: {e}")
            raise

    def get_last_modified_date(file_path):
        try:
            if os.path.exists(file_path):
                return datetime.fromtimestamp(os.path.getmtime(file_path))
            return None
        except Exception as e:
            print_log(f"Cannot access {file_path}: {e}")
            return None

    def get_days_ago(last_modified):
        if not last_modified:
            return ""
        days = (datetime.now() - last_modified).days
        return f"({days} days ago)"

    def is_workbook_open(excel, file_path):
        try:
            for wb in excel.Workbooks:
                if os.path.samefile(wb.FullName, file_path):
                    return True
            return False
        except Exception:
            return False

    def check_and_update_links():
        excel = setup_excel()
        threshold_date = datetime.now() - timedelta(days=CHECK_DAYS)
        total_updated = 0
        total_workbooks = 0

        print_log("=== Excel Link Update Started ===")
        print_log(f"Checking links modified within {CHECK_DAYS} days")

        summary_records = []

        try:
            total_workbooks = excel.Workbooks.Count
            if total_workbooks == 0:
                print_log("No open workbooks found")
                return

            workbook_index = 0
            for workbook in excel.Workbooks:
                workbook_index += 1
                workbook_name = workbook.FullName
                filename = os.path.basename(workbook_name)
                folderpath = os.path.dirname(workbook_name)
                print_log("")
                print_log("=" * 50)
                print_log(f"Scanning ({workbook_index}/{total_workbooks}): {filename}")

                try:
                    links = workbook.LinkSources(win32com.client.constants.xlExcelLinks)
                    if not links:
                        print_log("")
                        print_log(f"Action ({workbook_index}/{total_workbooks}): No external links found")
                        summary_records.append((folderpath, filename, "No external links found"))
                        continue

                    total_links = len(links)
                    print_log(f"  Found {total_links} external link(s)")
                    print_log("")
                    print_log("-" * 60)
                    updated_links = []

                    for link_index, link in enumerate(links, 1):
                        if link_index > 1:
                            print_log("-" * 60)
                        last_modified = get_last_modified_date(link)
                        last_modified_str = last_modified.strftime('%Y-%m-%d %H:%M:%S') if last_modified else "Not accessible"
                        days_ago = get_days_ago(last_modified)
                        status = ""
                        link_display = link if SHOW_FULL_PATH else os.path.basename(link)

                        summary_records.append((folderpath, filename, link))

                        if SHOW_LINK:
                            print_log(f"  Link ({link_index}/{total_links}): {link_display}")
                        if SHOW_LAST_MODIFIED:
                            print_log(f"  Last Modified: {last_modified_str} {days_ago}")

                        if last_modified and last_modified >= threshold_date:
                            if is_workbook_open(excel, link):
                                status = "Source file currently open. Update skipped (data refreshed in open workbook)."
                            else:
                                updated_links.append(link)
                                status = "Proceeding to update external link."
                        else:
                            status = f"No update needed (Source file not modified within {CHECK_DAYS} days)."

                        if SHOW_STATUS:
                            print_log(f"  Status: {status}")

                    if updated_links:
                        print_log("")
                        print_log(f"Action ({workbook_index}/{total_workbooks}): {filename}")
                        print_log("  Updating links...")
                        for link in updated_links:
                            try:
                                workbook.UpdateLink(Name=link, Type=win32com.client.constants.xlExcelLinks)
                                link_display = link if SHOW_FULL_PATH else os.path.basename(link)
                                print_log(f"    Updated: {link_display}")
                                total_updated += 1
                            except Exception as e:
                                link_display = link if SHOW_FULL_PATH else os.path.basename(link)
                                print_log(f"    Failed to update {link_display}: {e}")
                    else:
                        print_log("")
                        print_log(f"Action ({workbook_index}/{total_workbooks}): No links need updating")

                except Exception as e:
                    print_log(f"Error processing {filename}: {e}")

        except Exception as e:
            print_log(f"Process failed: {e}")

        finally:
            print_log("\n=== Excel Link Update Completed ===")
            print_log(f"Summary:")
            print_log(f"  Workbooks processed: {total_workbooks}")
            print_log(f"  Links updated: {total_updated}")
            save_scan_summary = options.get("SAVE_SCAN_SUMMARY", False)
            summary_dir = options.get("SUMMARY_DIR", r"D:\Pzone\Log")
            if save_scan_summary and summary_records:
                try:
                    import openpyxl
                    from openpyxl import Workbook
                    if not os.path.exists(summary_dir):
                        os.makedirs(summary_dir, exist_ok=True)
                    summary_file = os.path.join(
                        summary_dir,
                        f"excel_link_scan_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    )
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "External Linkage Scan"
                    ws.append(["Scanned file path", "Scanned file", "External Linkage"])
                    for rec in summary_records:
                        ws.append(list(rec))
                    wb.save(summary_file)
                    print_log(f"Scan summary excel saved: {summary_file}")
                except Exception as e:
                    print_log(f"Failed to save scan summary excel: {e}")
            pythoncom.CoUninitialize()
            if log_file_handler:
                log_file_handler.close()

    check_and_update_links()

if __name__ == "__main__":
    run_excel_link_update({}, print_func=None)