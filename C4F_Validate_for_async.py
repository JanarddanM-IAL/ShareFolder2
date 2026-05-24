import os
import re
import shutil
import traceback
import logging
from datetime import datetime
import fitz 
import pandas as pd
from PyPDF2 import PdfReader
from PyPDF2.errors import PdfReadError
from langdetect import detect
from langdetect.lang_detect_exception import LangDetectException
from typing import List, Dict, Any, Optional
from functools import lru_cache  # NEW: caching for speed
from dateutil import parser
from calendar import monthrange
import unicodedata
import numpy as np
import glob
from openpyxl import load_workbook


# Config
KEYMASTER_PATH = r"KeysMaster.xlsx" 
SAVE_RESULTS_IN_PDF_FOLDER = True  # save inside selected folder
OUTPUT_DIR = r"Output\\KeysMaster.xlsx"  # used only if SAVE_RESULTS_IN_PDF_FOLDER=False

# Regex behavior
GLOBAL_CASE_INSENSITIVE = True  # apply IGNORECASE to all patterns unless overridden with inline (?i)
REQUIRE_ALL_NONEMPTY_CHILDREN = True  # True: parent + ALL non-empty children must match
MIN_CHILD_MATCHES = 1  # used only if REQUIRE_ALL_NONEMPTY_CHILDREN=False

# Translation table for str.translate:
# - Map non-breaking spaces to ordinary spaces
# - Remove zero-width & BOM
# - Map various Unicode dashes to ASCII hyphen-minus
TRANS = {
    0x00A0: 0x20,  # NBSP -> space
    0x202F: 0x20,  # narrow NBSP -> space
    0x3000: 0x20,  # ideographic space -> space
    0x2000: 0x20, 0x2001: 0x20, 0x2002: 0x20, 0x2003: 0x20,
    0x2004: 0x20, 0x2005: 0x20, 0x2006: 0x20, 0x2007: 0x20,
    0x2008: 0x20, 0x2009: 0x20, 0x200A: 0x20, 0x205F: 0x20,
    0x200B: None,  # zero-width space -> remove
    0x2060: None,  # word joiner -> remove
    0xFEFF: None,  # BOM -> remove
    0x2010: ord('-'),  # hyphen -> '-'
    0x2011: ord('-'),  # non-breaking hyphen -> '-'
    0x2012: ord('-'),  # figure dash -> '-'
    0x2013: ord('-'),  # en dash -> '-'
    0x2014: ord('-'),  # em dash -> '-'
    0x2212: ord('-'),  # minus sign -> '-'
}


def build_record(
        name,
        pdf_path,
        inferred_type,
        res,
        created_date="",
        modified_date=""
    ):
        return {
            "pdf_name": name.replace("_AR", "").replace("_SR", "").replace(".pdf", ""),
            "pdf_path": pdf_path,
            "inferred_type": inferred_type,
            **res,
            "created_date": created_date,
            "modified_date": modified_date,
            "fiscalYearDate": res.get("fiscalYearDate", "")
        }

def worker(folder, stop_event=None):

    """
    Standalone worker function.

    Args:
        folder (str): Folder path containing PDFs and Portfolio Tracker.
        stop_event (threading.Event | None): Optional stop signal.

    Returns:
        tuple: (results, output_path)
            results: list of processed records
            output_path: generated Excel file path or None
    """

    results = []
    total = 0
    processed = 0
    output_path = None

    # -----------------------------
    # Logging setup
    # -----------------------------
    log_path = os.path.join(
        folder,
        f"pdf_processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    logger = logging.getLogger("pdf_worker_logger")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    def log_info(message):
        print(message)
        logger.info(message)

    def log_error(message, exc=None):
        print(message)
        if exc:
            logger.exception(message)
        else:
            logger.error(message)

    def is_stop_requested():
        return stop_event is not None and stop_event.is_set()

    

    def store_record_and_update_tracker(
        idx,
        record,
        dfPortfolioTracker,
        year="NA"
    ):
        nonlocal processed

        results.append(record)
        processed += 1

        dfPortfolioTracker.at[idx, "Created Date (DDMMYYYY)"] = record.get("created_date", "")
        dfPortfolioTracker.at[idx, "Modified Date (DDMMYYYY)"] = record.get("modified_date", "")
        dfPortfolioTracker.at[idx, "Fiscal Year Date (DDMMYYYY)"] = record.get("fiscalYearDate", "")
        dfPortfolioTracker.at[idx, "Validation"] = record.get("is_correct", False)
        # dfPortfolioTracker.at[idx, "Year"] = year
        dfPortfolioTracker.at[idx, "Year"] = "" if pd.isna(year) else str(year)
        dfPortfolioTracker.at[idx, "Remarks"] = record.get("status", "")

        log_info(f"Processed {processed}/{total}: {record.get('pdf_name', '')} | {record.get('status', '')}")

    try:
        language = ""

        # -----------------------------
        # Load KeyMaster rows
        # -----------------------------
        try:
            km_rows_all = load_keymaster_rows_by_type(KEYMASTER_PATH)
            log_info("KeyMaster loaded successfully.")
        except Exception as e:
            log_error(f"KeyMaster Error: {e}", e)
            return results, output_path

        # -----------------------------
        # Load Portfolio Tracker
        # -----------------------------
        try:
            # tracker_path = os.path.join(folder, "Carbon4finance_Portfolio Tracker.xlsx")

            # Get the single Excel file from folder
            excel_files = [
                f for f in glob.glob(os.path.join(folder, "*.xlsx"))
                if not os.path.basename(f).startswith("~$")
            ]

            if len(excel_files) == 0:
                raise FileNotFoundError("No Excel file found in the folder.")

            if len(excel_files) > 1:
                raise Exception("More than one Excel file found in the folder.")

            tracker_path = excel_files[0]

            # Load the Excel file
            xlsx = pd.ExcelFile(tracker_path, engine="openpyxl")

            # Check if "Tracker-IT" exists
            sheet_name = "Tracker-IT" if "Tracker-IT" in xlsx.sheet_names else xlsx.sheet_names[0]

            # Read the selected sheet
            dfPortfolioTracker = pd.read_excel(
                xlsx,
                sheet_name=sheet_name,
                header=0
            )

            for col in ["Year", "Validation", "Remarks", "Fiscal Year Date (DDMMYYYY)", "Created Date (DDMMYYYY)", "Modified Date (DDMMYYYY)"]:
                if col not in dfPortfolioTracker.columns:
                    dfPortfolioTracker[col] = ""
                dfPortfolioTracker[col] = dfPortfolioTracker[col].astype("object")

            required_cols = ["Report Name", "C4F Entity ID"]
            missing = [c for c in required_cols if c not in dfPortfolioTracker.columns]

            # print(dfPortfolioTracker.columns)
            if missing:
                raise KeyError(f"Missing columns in Tracker: {missing}")

            log_info("Portfolio Tracker loaded successfully.")

        except Exception as e:
            log_error(f"Portfolio Tracker Error: {e}", e)
            return results, output_path

        # -----------------------------
        # PDF set
        # -----------------------------
        pdf_files = {
            f.lower()
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        }

        # Count total PDFs expected to process
        for _, row in dfPortfolioTracker.iterrows():
            entity_name = str(row["Report Name"]).strip()

            if not entity_name:
                continue

            name_with_ext = (
                entity_name
                if entity_name.lower().endswith(".pdf")
                else f"{entity_name}.pdf"
            )

            if name_with_ext.lower() in pdf_files:
                total += 1

        log_info(f"Total PDFs found for processing: {total}")

        # -----------------------------
        # Main loop
        # -----------------------------
        for idx, row in dfPortfolioTracker.iterrows():

            if is_stop_requested():
                log_info("Stop requested. Processing stopped.")
                break

            entity_name = str(row["Report Name"]).strip()
            C4FEntityID = str(row["C4F Entity ID"]).strip()

            if not entity_name:
                continue

            name_with_ext = (
                entity_name
                if entity_name.lower().endswith(".pdf")
                else f"{entity_name}.pdf"
            )

            pdf_path = None

            if name_with_ext.lower() in pdf_files:
                pdf_path = os.path.join(folder, name_with_ext)

            if not pdf_path:
                dfPortfolioTracker.at[idx, "Remarks"] = "PDF not found"
                log_info(f"PDF not found: {entity_name}")
                continue

            res = {
                "fiscalYearDate": "",
                "is_correct": False,
                "matched_row_index": None,
                "matched_parent": "",
                "parent_pages": [],
                "matched_children": [],
                "children_pages_map": {},
                "children_required": "",
                "children_matched": "",
                "status": ""
            }

            name = os.path.basename(pdf_path)
            metadata = {}
            created_date = "Date not available"
            modified_date = "Date not available"
            inferred_type = None
            page_texts = []

            # -----------------------------
            # Metadata and language checks
            # -----------------------------
            try:
                inferred_type = parse_report_type_from_name(name)

                reader = PdfReader(pdf_path)

                if getattr(reader, "is_encrypted", False):
                    try:
                        reader.decrypt("")
                    except Exception:
                        pass

                metadata = reader.metadata or {}
                created_date = convert_pdf_date(metadata.get("/CreationDate"))
                modified_date = convert_pdf_date(metadata.get("/ModDate"))

                if len(reader.pages) < 10:
                    res["is_correct"] = False
                    res["status"] = "PDF pages is less than 10"

                    record = build_record(
                        name,
                        pdf_path,
                        inferred_type,
                        res,
                        created_date,
                        modified_date
                    )

                    store_record_and_update_tracker(idx, record, dfPortfolioTracker)
                    continue

                text, has_text = fast_extract_pdf_text(pdf_path)

                if has_text:
                    try:
                        language = detect(text)
                    except LangDetectException:
                        language = ""
                        res["status"] = "Language detection failed"
                        log_error(f"Language detection failed for: {name}")

                    if language != "en":
                        res["status"] = "Non-English Report"

                        record = build_record(
                            name,
                            pdf_path,
                            inferred_type,
                            res,
                            created_date,
                            modified_date
                        )

                        store_record_and_update_tracker(idx, record, dfPortfolioTracker)
                        continue
                else:
                    language = "No text found"
                    res["status"] = "PDF Contains Images Please check"

                    record = build_record(
                        name,
                        pdf_path,
                        inferred_type,
                        res,
                        created_date,
                        modified_date
                    )

                    store_record_and_update_tracker(idx, record, dfPortfolioTracker)
                    continue

            except PdfReadError as e:
                res["status"] = "Corrupted PDF - Unable to open"
                language = "Language detection skipped due to corrupted PDF"
                created_date = "Date not available"
                modified_date = "Date not available"

                record = build_record(
                    name,
                    pdf_path,
                    inferred_type,
                    res,
                    created_date,
                    modified_date
                )

                store_record_and_update_tracker(idx, record, dfPortfolioTracker)
                log_error(f"Corrupted PDF: {name}", e)
                continue

            except Exception as e:
                language = f"Language detection skipped: {e}"
                log_error(f"Metadata/language processing error for {name}: {e}", e)

            log_info(f"Processing {processed + 1} of {total}: {name}")

            # -----------------------------
            # Main PDF matching logic
            # -----------------------------
            try:
                page_texts = extract_text_per_page(pdf_path)

                if inferred_type == "AR" or inferred_type == "SR":
                    first_page_text = page_texts[0] if page_texts else ""

                    if first_page_text == "":
                        first_page_text = page_texts[1] if len(page_texts) > 1 else ""


                    is_wrong_type = re.search(
                        r"(Quarterly\s+Report|Interim\s+Report|Half[\s-]?Year"
                        r"|\b10-?Q\b|\b6-?K\b|\b8-?K\b"
                        r"|Immediate\s+Report|Proxy\s+Statement|Prospectus"
                        r"|\bSupplement(?:al|ary)?\b(?!\s+(?:ing\s+the\s+)?(?:Annual|Sustainability|ESG|Integrated)\s+Report)"
                        r"|(Annual|Sustainability|ESG|Integrated|Corporate)\s+Report\s+(19[0-9]{2}|20[01][0-9]|202[0-4])\b"   # "Annual Report 2024" and below
                        r"|\bFY\s?(19[0-9]{2}|20[01][0-9]|202[0-4])\b"                                                        # "FY2024", "FY 2023" and below
                        r"|(19[0-9]{2}|20[01][0-9]|202[0-4])\s+(Annual|Sustainability|ESG|Corporate|ENVIRONMENTAL)\b)",        # "2024 Corporate" and below
                        first_page_text,
                        re.IGNORECASE
                    )

                    if is_wrong_type:
                        res["is_correct"] = False
                        res["status"] = "Incorrect or Old report"

                        record = build_record(
                            name,
                            pdf_path,
                            inferred_type,
                            res,
                            created_date,
                            modified_date
                        )

                        store_record_and_update_tracker(idx, record, dfPortfolioTracker)
                        continue

                # Filter KeyMaster by type
                if inferred_type is None:
                    res = {
                        "is_correct": False,
                        "matched_row_index": None,
                        "matched_parent": "",
                        "parent_pages": [],
                        "matched_children": [],
                        "children_pages_map": {},
                        "children_required": "",
                        "children_matched": "",
                        "status": "",
                        "fiscalYearDate": ""
                    }
                else:
                    rows_subset = [
                        r for r in km_rows_all
                        if r["report_type"] == inferred_type
                    ]

                    if not rows_subset:
                        res = {
                            "is_correct": False,
                            "matched_row_index": None,
                            "matched_parent": "",
                            "parent_pages": [],
                            "matched_children": [],
                            "children_pages_map": {},
                            "children_required": "",
                            "children_matched": "",
                            "status": f"No KeyMaster rows for report type '{inferred_type}'.",
                            "fiscalYearDate": ""
                        }
                    else:
                        dateAR = ""

                        if inferred_type == "SR":
                            ar_entity = entity_name.replace("SR", "AR")
                            mask = dfPortfolioTracker["Report Name"].astype(str).str.strip().eq(ar_entity)
                            filtered_df = dfPortfolioTracker.loc[mask]

                            if (
                                not filtered_df.empty
                                and "Fiscal Year Date (DDMMYYYY)" in filtered_df.columns
                            ):
                                ser = (
                                    filtered_df["Fiscal Year Date (DDMMYYYY)"]
                                    .dropna()
                                    .astype(str)
                                    .str.strip()
                                    .str.replace(r"^\s*Date-\s*", "", regex=True)
                                )

                                if not ser.empty:
                                    dateAR = ser.iloc[0]

                        if res["status"] == "" and language == "en":
                            res = match_pdf_with_rows(
                                page_texts,
                                rows_subset,
                                dateAR,
                                pdf_path
                            )

                            # debug
                            # print(f"[ASSEMBLY] is_correct={res.get('is_correct')}")
                            # print(f"[ASSEMBLY] matched_children count={len(res.get('matched_children', []))}")
                            # print(f"[ASSEMBLY] children_pages_map keys:")
                            # for k, v in res.get('children_pages_map', {}).items():
                            #     print(f"  key={k[:60]}... → pages={v}")


                record = build_record(
                    name,
                    pdf_path,
                    inferred_type,
                    res,
                    created_date,
                    modified_date
                )

            except Exception as e:
                record = {
                    "pdf_name": name.replace("_AR", "").replace("_SR", ""),
                    "pdf_path": pdf_path,
                    "inferred_type": inferred_type,
                    "is_correct": False,
                    "matched_row_index": None,
                    "matched_parent": "",
                    "parent_pages": [],
                    "matched_children": [],
                    "children_pages_map": {},
                    "children_required": "",
                    "children_matched": "",
                    "status": f"[ERROR] {e}",
                    "created_date": "",
                    "modified_date": "",
                    "fiscalYearDate": ""
                }

                log_error(f"PDF matching error for {name}: {e}", e)

            # -----------------------------
            # Fiscal year and company check
            # -----------------------------
            date_obj = None
            fy_val = str(record.get("fiscalYearDate", "") or "")
            stat = record.get("is_correct")
            year = "NA"

            if (
                fy_val
                and "Not Matched with AR Report" not in fy_val
                and "Date-" in fy_val
            ):
                raw_date = fy_val.replace("Date-", "").strip()
                s = str(raw_date).strip()

                for fmt in (
                    "%d/%m/%Y",
                    "%m/%d/%Y",
                    "%b. %d, %Y",
                    "%b %d, %Y",
                    "%d %B %Y",
                    "%B %d, %Y",
                    "%d%B%Y"
                ):
                    try:
                        date_obj = datetime.strptime(s, fmt)
                        break
                    except ValueError:
                        continue

                if date_obj is None:
                    normalized = "".join(ch for ch in s if ch.isalnum())

                    for fmt in ("%d%m%Y", "%m%d%Y", "%Y%m%d"):
                        try:
                            date_obj = datetime.strptime(normalized, fmt)
                            break
                        except ValueError:
                            continue

                if date_obj:
                    fiscalYearDate1 = date_obj.strftime("%d%m%Y")

                    if len(fiscalYearDate1) == 8 and fiscalYearDate1.isdigit():
                        year = int(fiscalYearDate1[-4:])
                    else:
                        year = "NA"

            elif re.match(r"^\d{4}$", str(fy_val)):
                year = fy_val

            # Company Name check
            if stat is True:
                company_name = " ".join(
                    entity_name.lower()
                    .replace("_ar", "")
                    .replace("_sr", "")
                    .replace(".pdf", "")
                    .replace("corp.", "")
                    .replace("corp", "")
                    .replace("ltd.", "")
                    .replace("the ", "")
                    .replace("inc.", "")
                    .replace(" ", "")
                    .replace(",", "")
                    .replace(".", "")
                    .replace("&", "")
                    .replace("'", "")
                    .split()[:2]
                )

                # raw_first_10 = " ".join(page_texts[:10])

                pages_to_search = page_texts[:10] #raw_first_10 page
                if len(page_texts) > 10:
                    pages_to_search = pages_to_search + page_texts[-2:] #add last 2 page text also

                raw_first_10 = " ".join(pages_to_search)


                first_10_pages_text = (
                    normalize_pdf_text(raw_first_10)
                    .lower()
                    .replace("_ar", "")
                    .replace("_sr", "")
                    .replace(".pdf", "")
                    .replace("corp.", "")
                    .replace("corp", "")
                    .replace("ltd.", "")
                    .replace("the ", "")
                    .replace("inc.", "")
                    .replace(" ", "")
                    .replace(",", "")
                    .replace(".", "")
                    .replace("&", "")
                    .replace("'", "")
                    .replace("\n", "")
                )

                if company_name and company_name.lower() not in first_10_pages_text.lower():
                    record["is_correct"] = True
                    record["status"] = "Parent matched but Company name Mismatch"
                else:
                    record["is_correct"] = True
                    record["status"] = "Parent & Company Name Matched"

            if date_obj is not None:
                if len(fiscalYearDate1) == 8 and fiscalYearDate1.isdigit():
                    year = int(fiscalYearDate1[-4:])

            store_record_and_update_tracker(idx, record, dfPortfolioTracker, year)

        # -----------------------------
        # Export results
        # -----------------------------
        if results:
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")

                out_dir = folder if SAVE_RESULTS_IN_PDF_FOLDER else OUTPUT_DIR

                if not SAVE_RESULTS_IN_PDF_FOLDER:
                    os.makedirs(out_dir, exist_ok=True)

                # output_path = os.path.join(
                #     out_dir,
                #     f"TypeFiltered_Tagging_Results_{ts}.xlsx"
                # )

                output_path =tracker_path

                df = pd.DataFrame(results)

                df["Correct"] = df["is_correct"].map(
                    lambda x: "Yes" if x else "No"
                )

                df["ParentPages"] = df["parent_pages"].map(
                    lambda lst: ",".join(map(str, lst)) if lst else ""
                )

                df["ChildrenPages"] = df["children_pages_map"].map(
                    format_children_pages
                )

                df_details = df[[
                    "pdf_name",
                    "pdf_path",
                    "inferred_type",
                    "Correct",
                    "matched_row_index",
                    "matched_parent",
                    "children_matched",
                    "children_required",
                    "ParentPages",
                    "ChildrenPages",
                    "status",
                    "created_date",
                    "modified_date",
                    "fiscalYearDate"
                ]].rename(columns={
                    "pdf_name": "PDF Name",
                    "pdf_path": "PDF Path",
                    "inferred_type": "Report Type (from filename)",
                    "matched_row_index": "Matched Row Index",
                    "matched_parent": "Parent Regex",
                    "children_matched": "Children Matched",
                    "children_required": "Children Required",
                    "ParentPages": "Parent Pages",
                    "ChildrenPages": "Children Pages",
                    "status": "Status",
                    "created_date": "Created Date (DDMMYYYY)",
                    "modified_date": "Modified Date (DDMMYYYY)",
                    "fiscalYearDate": "Fiscal Year Date (DDMMYYYY)"
                })

                fy_col = "Fiscal Year Date (DDMMYYYY)"
                fy_col_ModifiedDate = "Modified Date (DDMMYYYY)"
                remarks_col = "Remarks"

                fy_str = dfPortfolioTracker[fy_col].astype("string")
                remarks_str = dfPortfolioTracker[remarks_col].astype("string")

                # Normalize Modified Date
                MOD_COL = "Modified Date (DDMMYYYY)"
                dfPortfolioTracker[MOD_COL] = dfPortfolioTracker[MOD_COL].apply(convert_to_ddmmyyyy)

                mod_series = dfPortfolioTracker[MOD_COL].astype(str)
                mod_series_clean = mod_series.str.replace(r"\D", "", regex=True)

                mod_year = pd.to_datetime(
                    mod_series_clean,
                    format="%d%m%Y",
                    errors="coerce"
                ).dt.year

                if "CorrectedByModYear" not in dfPortfolioTracker.columns:
                    dfPortfolioTracker["CorrectedByModYear"] = False

                mask_last_year = remarks_str.str.contains(
                    "This is last year report",
                    case=False,
                    na=False
                )

                mask_lt_2024 = remarks_str.str.contains(
                    "Year is less than 2024",
                    case=False,
                    na=False
                )

                mask_target = mask_last_year | mask_lt_2024
                mask_mod_gt24 = mod_year > 2024
                mask_fix = mask_target & mask_mod_gt24

                dfPortfolioTracker.loc[mask_fix, "Validation"] = "TRUE"
                dfPortfolioTracker.loc[mask_fix, "Remarks"] = (
                    "Modified year is > 2024.Please check Fiscal Year Date"
                )
                dfPortfolioTracker.loc[mask_fix, "CorrectedByModYear"] = True

                # Normalize fiscal year date except year-only values
                fy_str = dfPortfolioTracker[fy_col].astype("string")

                mask_not_year_only = ~fy_str.str.match(r"^\d{4}$", na=False)

                dfPortfolioTracker.loc[mask_not_year_only, fy_col] = (
                    dfPortfolioTracker.loc[mask_not_year_only, fy_col]
                    .apply(convert_to_ddmmyyyy)
                )

                fy_clean = dfPortfolioTracker[fy_col].astype("string")

                mask_ddmmyyyy = fy_clean.str.match(r"^\d{8}$", na=False)
                mask_year_only = fy_clean.str.match(r"^\d{4}$", na=False)

                year_series = pd.Series(
                    pd.NA,
                    index=dfPortfolioTracker.index,
                    dtype="object"
                )

                year_series.loc[mask_ddmmyyyy] = fy_clean.loc[mask_ddmmyyyy].str[-4:]
                year_series.loc[mask_year_only] = fy_clean.loc[mask_year_only]

                year_numeric = pd.to_numeric(year_series, errors="coerce")

                dfPortfolioTracker["Year"] = dfPortfolioTracker["Year"].mask(
                    dfPortfolioTracker["Year"].isna()
                    | dfPortfolioTracker["Year"].astype("string").str.strip().str.upper().eq("NA"),
                    year_numeric
                )

                # Validation
                if fy_str.str.match(r"^\d{4}$").all():
                    tmp = dfPortfolioTracker[[fy_col]].copy()
                    tmp["Validation_tmp"] = ""
                    tmp["NewRemark"] = ""
                else:
                    tmp = dfPortfolioTracker[fy_col].apply(
                        lambda x: pd.Series(
                            validate_year(x),
                            index=["Validation_tmp", "NewRemark"]
                        )
                    )

                new_remark_str = tmp["NewRemark"].astype("string")
                mask_update_remarks = (
                    new_remark_str.notna()
                    & new_remark_str.str.strip().ne("")
                )

                dfPortfolioTracker.loc[mask_update_remarks, remarks_col] = (
                    new_remark_str[mask_update_remarks]
                )

                remarks_str = dfPortfolioTracker[remarks_col].astype("string")
                fy_str = dfPortfolioTracker[fy_col].astype("string")

                # Parent matched and Fiscal Year Date blank
                mask_blank = (
                    remarks_str.str.contains("Parent matched", case=False, na=False)
                    & fy_str.fillna("").str.strip().eq("")
                )

                dfPortfolioTracker.loc[mask_blank, remarks_col] = "Fiscal Year Date not found"
                dfPortfolioTracker.loc[mask_blank, "Validation"] = "FALSE"

                # Parent matched and year-only
                mod_series = dfPortfolioTracker[MOD_COL].astype("string")
                mod_series_clean = mod_series.str.replace(r"\D", "", regex=True)

                mod_year = pd.to_datetime(
                    mod_series_clean,
                    format="%d%m%Y",
                    errors="coerce"
                ).dt.year

                remarks_str = dfPortfolioTracker[remarks_col].astype("string")
                fy_str = dfPortfolioTracker[fy_col].astype("string")

                mask_year_only = (
                    remarks_str.str.contains("Parent matched", case=False, na=False)
                    & fy_str.str.match(r"^\d{4}$", na=False)
                )

                dfPortfolioTracker.loc[mask_year_only, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_year_only, remarks_col] = (
                    "Parent matched - Only Fiscal Year found in report"
                )

                mask_year_2025 = mask_year_only & mod_year.notna() & mod_year.gt(2024)

                dfPortfolioTracker.loc[mask_year_2025, "Validation"] = "TRUE"
                dfPortfolioTracker.loc[mask_year_2025, "Remarks"] = (
                    "Parent matched - Modified year is > 2024"
                )

                remarks_str = dfPortfolioTracker[remarks_col].astype("string")

                # Other conditions
                mask_year_2024 = remarks_str.str.contains(
                    "Year is less than 2024",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[
                    mask_year_2024 & ~dfPortfolioTracker["CorrectedByModYear"],
                    "Validation"
                ] = "FALSE"

                mask_corrupted = remarks_str.str.contains(
                    "Failed to open file|Corrupted PDF",
                    case=False,
                    na=False,
                    regex=True
                )

                dfPortfolioTracker.loc[mask_corrupted, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_corrupted, "Remarks"] = (
                    "Corrupted PDF - Unable to open"
                )

                mask_language = remarks_str.str.contains(
                    "Non-English Report",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_language, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_language, "Remarks"] = (
                    "PDF Language is not english"
                )

                mask_small_pdf = remarks_str.str.contains(
                    "PDF pages is less than 10",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_small_pdf, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_small_pdf, "Remarks"] = (
                    "PDF pages is less than 10"
                )

                mask_fy_report = remarks_str.str.contains(
                    "Proper Financial Statement key word not available within the 3 pages",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_fy_report, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_fy_report, "Remarks"] = (
                    "Proper Financial Statement key word not available within the 3 pages"
                )

                mask_10q = remarks_str.str.contains(
                    "Quarterly report detected",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_10q, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_10q, "Remarks"] = (
                    "Quarterly report detected"
                )

                # Last year report override
                mod_series = dfPortfolioTracker[MOD_COL].astype("string")
                mod_series_clean = mod_series.str.replace(r"\D", "", regex=True)

                mod_year = pd.to_datetime(
                    mod_series_clean,
                    format="%d%m%Y",
                    errors="coerce"
                ).dt.year

                remarks_str = dfPortfolioTracker[remarks_col].astype("string")

                mask_last_year_report = remarks_str.str.contains(
                    "This is last year report",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_last_year_report, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_last_year_report, remarks_col] = (
                    "This is last year report."
                )

                mask_last_year_report_fix = (
                    mask_last_year_report
                    & mod_year.notna()
                    & mod_year.gt(2024)
                )

                dfPortfolioTracker.loc[mask_last_year_report_fix, "Validation"] = "TRUE"
                dfPortfolioTracker.loc[mask_last_year_report_fix, remarks_col] = (
                    "This is last year report - Overridden because Modified year is > 2024"
                )

                # Created Date check
                CRE_COL = "Created Date (DDMMYYYY)"
                dfPortfolioTracker[CRE_COL] = dfPortfolioTracker[CRE_COL].apply(convert_to_ddmmyyyy)

                cre_series = dfPortfolioTracker[CRE_COL].astype("string")
                cre_series_clean = cre_series.str.replace(r"\D", "", regex=True)

                cre_year = pd.to_datetime(
                    cre_series_clean,
                    format="%d%m%Y",
                    errors="coerce"
                ).dt.year

                val_col = "Validation"

                dfPortfolioTracker[val_col] = (
                    dfPortfolioTracker[val_col]
                    .map(
                        lambda x:
                        "TRUE" if str(x).strip().lower() in {"true", "yes", "1"}
                        else "FALSE" if str(x).strip().lower() in {"false", "no", "0"}
                        else str(x)
                    )
                    .astype("string")
                    .str.upper()
                )

                mask_valid_true = dfPortfolioTracker[val_col].eq("TRUE")
                mask_cre_fix = mask_valid_true & cre_year.notna() & cre_year.eq(2024)

                dfPortfolioTracker.loc[mask_cre_fix, val_col] = "TRUE"
                dfPortfolioTracker.loc[mask_cre_fix, remarks_col] = (
                    "This is last year report - Overridden because Created Date is < 2025"
                )

                remarks_str = dfPortfolioTracker[remarks_col].astype("string")

                mask_after_page_no_4 = remarks_str.str.contains(
                    "Parent Key Word ater Page no. 4",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_after_page_no_4, "Validation"] = "FALSE"
                dfPortfolioTracker.loc[mask_after_page_no_4, "Remarks"] = (
                    "Parent Key Word ater Page no. 4"
                )

                mask_company_name = remarks_str.str.contains(
                    "Parent matched but Company name Mismatch",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_company_name, "Validation"] = "TRUE"
                dfPortfolioTracker.loc[mask_company_name, "Remarks"] = (
                    "Parent matched but Company name Mismatch"
                )

                mask_mod_date = remarks_str.str.contains(
                    "Modified year is > 2024",
                    case=False,
                    na=False
                )

                dfPortfolioTracker.loc[mask_mod_date, "Validation"] = "TRUE"
                dfPortfolioTracker.loc[mask_mod_date, "Remarks"] = (
                    "Modified year is > 2024.Please check Fiscal Year Date"
                )

                # Generate summary
                report_summary = generate_summary(dfPortfolioTracker)

                # Copy corrected PDFs
                corrected_folder = os.path.join(folder, "Corrected Reports")
                os.makedirs(corrected_folder, exist_ok=True)

                for _, row in dfPortfolioTracker.iterrows():
                    if str(row.get("Validation", "")).strip().upper() == "TRUE":
                        original_name = str(row.get("Report Name", "")).strip()

                        actual_name = (
                            str(row.get("C4F Entity ID", "")).strip()
                            + "_"
                            + str(row.get("Year", "")).strip()
                            + "_"
                            + str(row.get("Report Type", "")).strip()
                            + "_"
                            + str(row.get("Entity Name", "")).strip()
                            .replace("_AR", "")
                            .replace("_SR", "")
                            .replace(".pdf", "")
                        )

                        if not original_name or not actual_name:
                            continue

                        original_pdf = (
                            original_name
                            if original_name.lower().endswith(".pdf")
                            else f"{original_name}.pdf"
                        )

                        source_path = os.path.join(folder, original_pdf)
                        target_path = os.path.join(corrected_folder, f"{actual_name}.pdf")

                        if os.path.isfile(source_path):
                            shutil.copy2(source_path, target_path)

                with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                    report_summary.to_excel(writer, index=False, sheet_name="Summary")
                    df_details.to_excel(writer, index=False, sheet_name="Details")
                    dfPortfolioTracker.to_excel(writer, index=False, sheet_name="Tracker-IT")

                    workbook = writer.book
                    workbook["Details"].sheet_state = "hidden"

                log_info(
                    f"Done. Processed {processed}/{total} PDFs. "
                    f"Results saved to: {output_path}"
                )

            except Exception as e:
                log_error(f"Export Error: {e}", e)

        else:
            log_info("No results found. Nothing exported.")

        log_info(f"Log file saved to: {log_path}")
        return results, output_path

    except Exception as e:
        tb = traceback.format_exc()
        log_error(f"Unexpected Error: {e}\n\n{tb}", e)
        return results, output_path
    
def _norm(s: str) -> str:
    """Normalize column name: lowercase, remove non-alphanumerics."""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def load_keymaster_rows_by_type(xlsx_path: str):
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"KeyMaster not found: {xlsx_path}")
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    if df.empty:
        raise ValueError("KeyMaster Excel is empty.")
    # Detect columns robustly
    cols_norm = { _norm(c): c for c in df.columns }
    # Report Type
    report_col = None
    for norm, orig in cols_norm.items():
        if norm in {"reporttype", "reporttypes"}:
            report_col = orig
            break
    if report_col is None:
        raise ValueError("Column 'Report Type' not found in KeyMaster.")
    # Parent
    parent_col = None
    for norm, orig in cols_norm.items():
        if norm == "parentkeywords":
            parent_col = orig
            break
    if parent_col is None:
        # fallback variants
        for norm, orig in cols_norm.items():
            if norm in {"parentkeyword", "parentkeys", "parentkey", "parent"}:
                parent_col = orig
                break
    if parent_col is None:
        raise ValueError("Column 'Parent Key Words' not found in KeyMaster.")
    # Children (1..6)
    child_cols = []
    for i in range(1, 7):
        target_norm = f"childkeyword{i}"
        for norm, orig in cols_norm.items():
            if norm == target_norm:
                child_cols.append(orig)
                break
    rows = []
    for ridx, row in df.iterrows():
        report_type_raw = row.get(report_col, "")
        if pd.isna(report_type_raw):
            continue
        report_type = str(report_type_raw).strip().upper()
        if report_type not in {"AR", "SR"}:
            continue
        parent = str(row.get(parent_col, "")).strip()
        if not parent:
            continue
        children = []
        for c in child_cols:
            val = row.get(c, None)
            if pd.isna(val):
                continue
            sval = str(val).strip()
            if sval:
                children.append(sval)
        rows.append({
            "row_index": ridx + 1,
            "report_type": report_type,
            "parent": parent,
            "children": children
        })
    if not rows:
        raise ValueError("No usable rows found in KeyMaster (check 'Report Type' and 'Parent Key Words').")
    return rows

def parse_report_type_from_name(filename: str) -> str | None:
    base = os.path.splitext(os.path.basename(filename))[0].strip()
    low = base.lower()
    if low.endswith("_ar"):
        return "AR"
    if low.endswith("_sr"):
        return "SR"
    return None

def convert_pdf_date(pdf_date):
    if not pdf_date or not isinstance(pdf_date, str):
        return "Date not available"
    match = re.match(r"D:(\d{8})", pdf_date)
    if match:
        try:
            dt = datetime.strptime(match.group(1), "%Y%m%d")
            return dt.strftime("%d%m%Y")
        except ValueError:
            return "Invalid Date format"
    return "Date not available"

def fast_extract_pdf_text(pdf_path: str) -> tuple[str, bool]:
    """Very fast whole-PDF extraction using PyMuPDF only. Returns (full_text, has_text)."""
    try:
        doc = fitz.open(pdf_path)
        if doc.is_encrypted:
            try:
                doc.authenticate("")
            except Exception:
                return "", False
        texts = []
        for page in doc:
            t = page.get_text("text")
            if t:
                texts.append(t)
        full = "\n".join(texts).strip()
        return full, bool(full)
    except Exception:
        return "", False

def extract_text_per_page(pdf_path: str) -> list[str]:
    """Ultra-fast page-by-page extraction using a single PyMuPDF open.
    Logic identical (list of per-page strings), just faster and safer.
    """
    try:
        doc = fitz.open(pdf_path)
        if doc.is_encrypted:
            try:
                doc.authenticate("")
            except Exception:
                return []
        return [page.get_text("text") or "" for page in doc]
    except Exception:
        return []
    

# core validation engine
def match_pdf_with_rows(page_texts: List[str], rows_subset: List[dict], dateAR: str, pdf_path: str) -> Dict[str, Any]:
    # Safe defaults if globals are not defined elsewhere
    require_all = globals().get("REQUIRE_ALL_NONEMPTY_CHILDREN", False)
    min_child_matches = globals().get("MIN_CHILD_MATCHES", 1)
    full_text = "\n".join(page_texts)
    # --- Safe defaults used in the final return (prevents UnboundLocalError) ---
    parent_pages: List[int] = []  # NEW
    children_required: int = 0
    children_matched: int = 0  # NEW
    children_pages_map: Dict[str, List[int]] = {}  # NEW
    matched_children: List[str] = []  # NEW
    fiscalYearDate: Optional[str] = ""  # NEW
    status: str = "No row satisfied for this report type."  # NEW
    status1: str = ""  # NEW
    # Co-term lookahead maps (page must contain all lookahead terms)
    # SAME_PAGE_CO_TERMS_MAP_BS = {
    #     1: (
    #         r"(?i)"
    #         r"(?=.*\b(?:Total\s+Assets|Assets|Net\s+assets)\b)"
    #         r"(?=.*\b(?:Current\s+Assets|Total\s+current\s+assets|comprehensive\s+income|Fixed\s+Assets|Total\s+Financial\s+Assets)\b)"
    #         r"(?=.*\b(?:Total\s+equity|liabilities\s+and\s+equity|capital\s+and\s+liabilities|equity\s+and\s+liabilities|Total\s+liabilities|Other\s+current\s+assets|Fixed\s+assets|TOTAL\s+EQUITY\s+AND\s+LIABILITIES)\b)"
    #     )
    # }
    SAME_PAGE_CO_TERMS_MAP_BS = {
        1: (
            r"(?i)"
            r"(?=.*\b(?:Total\s+Assets|Assets|Net\s+assets|ASSETS)\b)"
            r"(?=.*\b(?:Current\s+Assets|Total\s+current\s+assets|"
            r"comprehensive\s+income|Fixed\s+Assets|Total\s+Financial\s+Assets|"
            r"Intangible\s+assets|Investments|Receivables|"    # ← Allianz/German structure
            r"Tangible\s+assets)\b)"
            r"(?=.*\b(?:Total\s+equity|liabilities\s+and\s+equity|"
            r"capital\s+and\s+liabilities|equity\s+and\s+liabilities|"
            r"Total\s+liabilities|TOTAL\s+EQUITY\s+AND\s+LIABILITIES|"
            r"Total\s+equity\s+and\s+liabilities|"            # ← Allianz exact phrase
            r"Shareholders.*equity|EQUITY\s+AND\s+LIABILITIES)\b)"  # ← Allianz
        )
    }

    # SAME_PAGE_CO_TERMS_MAP_PL = {
    #     2: (
    #         r"(?is)"  # i = ignorecase, s = dotall so .* spans lines
    #         r"(?=.*\b(?:Revenue|Sales|Operating\s+income|Operating\s+profit|EXPENDITURE|Other(?:s)?\s+income|Net\s+sales|operating\s+profit\s+or\s+loss|Interest\s+income)\b)"
    #         r"(?=.*\b(?:Profit\s+before\s+tax|Profit\s+after\s+tax(?:es)?|cash\s+flows\s+as\s+operating\s+activities|"
    #         r"Income\s+before\s+tax|Net\s+income|Profit\s+for\s+the\s+year|Total\s+income\s+tax(?:es)?|"
    #         r"Profit/\(loss\)\s+for\s+the|Profit/\(loss\)\s+before\s+taxation|Net\s+profit|Operating\s+profit|EARNINGS\s+PER\s+SHARE)\b)"
    #     )
    # }
    SAME_PAGE_CO_TERMS_MAP_PL = {
        2: (
            r"(?is)"
            # Group 1: Top line / revenue terms
            r"(?=.*\b(?:Revenue|Sales|Operating\s+income|Operating\s+profit|"
            r"EXPENDITURE|Other(?:s)?\s+income|Net\s+sales|operating\s+profit\s+or\s+loss|"
            r"Interest\s+income|"
            r"Insurance\s+income|"              # ← NN Group
            r"Investment\s+income|"             # ← NN Group
            r"Net\s+insurance\s+result|"        # ← NN Group
            r"Net\s+investment\s+result|"       # ← NN Group
            r"Net\s+interest\s+income|"         # ← banks
            r"Fee\s+and\s+commission\s+result)\b)"  # ← NN Group

            # Group 2: Bottom line / profit terms
            r"(?=.*\b(?:Profit\s+before\s+tax|Profit\s+after\s+tax(?:es)?|"
            r"cash\s+flows\s+as\s+operating\s+activities|"
            r"Income\s+before\s+tax|Net\s+income|Profit\s+for\s+the\s+year|"
            r"Total\s+income\s+tax(?:es)?|Profit/\(loss\)\s+for\s+the|"
            r"Profit/\(loss\)\s+before\s+taxation|Net\s+profit|Operating\s+profit|"
            r"EARNINGS\s+PER\s+SHARE|"
            r"Earnings\s+per\s+(?:ordinary\s+)?share|"  # ← NN Group ("ordinary" variant)
            r"Result\s+before\s+tax|"                   # ← NN Group
            r"Net\s+result)\b)"                          # ← NN Group
        )
    }


    SAME_PAGE_CO_TERMS_MAP_CF = {
        3: (
            r"(?is)"
            r"(?=.*\b(?:Cash\s+and\s+cash\s+equivalents|Net\s+cash\s+and\s+cash\s+equivalents|NET\s+CASH\s+FLOW|Net\s+cash|Change\s+in\s+net\s+cash\s+position|"
            r"Net\s+cash\s+generated|Net\s+cash\s+flows|Net\s+income|Income\s+taxes\s+paid|operating\s+profit\s+or\s+loss)\b)"
            r"(?=.*\b(?:investing\s+activities|Working\s+capital\s+adjustments|financing\s+activities|operating\s+activities|operating\s+cash\s+flows|"
            r"interest\s+and\s+dividend\s+cash\s+flows\s+as\s+operating\s+activities)\b)"
        )
    }
    for r in rows_subset:
        fiscalYearDate = ""  # ensure it's always defined
        # --- Parent regex ---
        try:
            rx_parent = compile_regex(r["parent"])
        except re.error as e:
            print(f"[WARN] Invalid parent regex (row {r.get('row_index')}): {r.get('parent')}; {e}")
            continue
        # Quick parent presence check
        if not rx_parent.search(full_text):
            status = f"No parent match for row {r.get('row_index')}"
            continue
        parent_pages = pages_for_regex(page_texts, rx_parent)
        
        # debug Parent Key Words regex 
        # print("_"*50)
        # print(rx_parent)
        # print("First parent page:", min(parent_pages) if parent_pages else "NA")
        
        
        if parent_pages and min(parent_pages) > 10:
            status1: str = "Parent found but after the page 10."
            # print("XO"*10)
            continue
        else:
            children_required = len(r["children"])
            children_pages_map = {}
            matched_children = []
            children_pages_map: Dict[str, List[int]] = {}
            matched_children: List[str] = []
            # SR special path when there are no children
            if r.get("report_type") == "SR":
                try:
                    rx_dateAR = compile_regex(dateAR) if dateAR else None
                    dates_pages = pages_for_regex(page_texts, rx_dateAR) if rx_dateAR else []
                    if len(dates_pages) > 0:
                        fiscalYearDate = "Date-" + dateAR
                    else:
                        ctx_dt_set, ctx_dt_to_raw = extract_balance_sheet_context_dates(page_texts)
                        if ctx_dt_set:
                            chosen = get_valid_max(ctx_dt_set)
                            raw = next(iter(ctx_dt_to_raw[chosen]))
                            fiscalYearDate = raw + " - Not Matched with AR Report"
                        elif dateAR:
                            try:
                                parsed_date = parser.parse(dateAR)
                                srYear = parsed_date.year
                                srMonth = parsed_date.strftime("%B")
                                date_pattern = rf"\b{srMonth}\b.*\b{srYear}\b|\b{srYear}\b.*\b{srMonth}\b"
                                rx_dateAR = compile_regex(date_pattern)
                                dates_pages = pages_for_regex(page_texts, rx_dateAR) if rx_dateAR else []
                                if len(dates_pages) > 0:
                                    fiscalYearDate = "Date-" + dateAR
                                else:
                                    date_patternYear = rf"\b{srYear}\b"
                                    rx_dateAR = compile_regex(date_patternYear)
                                    dates_pages = pages_for_regex(page_texts, rx_dateAR) if rx_dateAR else []
                                    if len(dates_pages) > 0:
                                        fiscalYearDate = dateAR
                            except Exception as e:
                                return f"Error parsing date: {e}"
                        else:
                            # If still empty, check for keywords like "Published", "Issued", etc.
                            if fiscalYearDate == "":
                                raw_end, end_date = extract_sr_fiscal_end_date(page_texts)
                                if end_date:
                                    # keep raw phrase for UI, your post-processing will convert "Date-..." to DDMMYYYY
                                    fiscalYearDate = "Date-" + end_date.strftime("%d%m%Y")
                                else:
                                    keywords = ["Published", "Issued", "Reporting Period", "Report Date", "Updated", "Effective","Time Frame","Fiscal Year Ended"]
                                    combined_text = " ".join(page_texts[:10])
                                    if any(kw in combined_text for kw in keywords):
                                        all_dt_set, all_dt_to_raw = extract_dates_from_text(combined_text.replace("Notes ", ""))
                                        if all_dt_set:
                                            march_dec_dates = [d for d in all_dt_set if d.month in (3, 12)]
                                            if march_dec_dates:
                                                chosen = get_valid_max(march_dec_dates)
                                                raw = next(iter(all_dt_to_raw[chosen]))
                                                fiscalYearDate = "Date-" + raw
                                    if fiscalYearDate == '':
                                        fiscalYearDate = None  # initialize
                                        rx_year = compile_regex(r"\b(?:20)\d{2}\b")
                                        rx_year1 = compile_regex(r"\b(?:20)\d{2}\s*-\s*\d{2}\b")
                                        best_year = None
                                        pages_to_check = range(min(3, len(page_texts)))  # page 1 and 2 only
                                        for p in pages_to_check:
                                            match = rx_year1.search(page_texts[p])
                                            if match:
                                                start_year = int(match.group(0)[0:4])
                                                end_suffix = int(match.group(0)[-2:])  # '24'
                                                end_year = (start_year // 100) * 100 + end_suffix  # 2024
                                                fiscalYearDate = str(end_year)
                                                break
                                            else:
                                                years_on_page = [int(y) for y in rx_year.findall(page_texts[p])]
                                                if years_on_page:
                                                    best_year = max(years_on_page)
                                                    fiscalYearDate = str(best_year)
                                                    break
                except re.error as e:
                    print(f"[WARN] Invalid dateAR regex '{dateAR}': {e}")
            for idx, child_pat in enumerate(r["children"], start=1):
                try:
                    rx_child = compile_regex(child_pat)
                except re.error as e:
                    print(f"[WARN] Invalid child regex (row {r.get('row_index')}): {child_pat}; {e}")
                    children_pages_map[child_pat] = []
                    continue
                child_pages = pages_for_regex(page_texts, rx_child)

                # debug Balance Sheet Table: ; Income statement: ; Cash Flow Table: 
                # print(f"[DEBUG] idx={idx}")
                # print(f"[DEBUG] child_pat:\n{child_pat}")
                # print(f"[DEBUG] child_pages={child_pages}")


                # if idx == 1:
                    # print("=" * 100)
                    # print("[PAGE SCAN] Scanning pages 15-48 for financial statement headings...")
                    # for page_num in range(15, 49):
                    #     if page_num <= len(page_texts):
                    #         lines = page_texts[page_num - 1].strip().split("\n")
                    #         for line in lines:
                    #             line = line.strip()
                    #             if not line:
                    #                 continue
                    #             keywords = ["balance sheet", "income statement", "profit and loss",
                    #                         "cash flow", "financial position", "statement of income",
                    #                         "statement of profit", "statement of cash"]
                    #             if any(kw in line.lower() for kw in keywords):
                    #                 print(f"  Page {page_num}: '{line}'")
                    #                 break  # only first matching line per page

                    # print(f"[PAGE 14 RAW]: {repr(page_texts[14][:1500])}")
                    # print(f"[PAGE 15 RAW]: {repr(page_texts[15][:1500])}")
                    # print(f"[PAGE 16 RAW]: {repr(page_texts[16][:1500])}")
                    # print("=" * 100)


                # debug
                # if ("balance" in child_pat.lower()  and "sheet" in child_pat.lower() ) or ("financial" in child_pat.lower()  and "position" in child_pat.lower()):
                #     print(f"[BS] child_pat:\n{child_pat}")
                #     print(f"[BS] child_pages={child_pages}")
                    
                    # Search nearby pages around where IS was found (212) and CF (210)
                    # BS is usually just before these, so check pages 205-215
                    # for page_num in range(205, 215):  #page number is for a specific file, it will be changed
                    #     if page_num <= len(page_texts):
                    #         page_sample = page_texts[page_num - 1][:500]
                    #         if page_sample.strip():
                    #             print(f"[BS] Page {page_num} sample:\n{page_sample}")
                    #             print("-" * 50)


                # debug for profit loass and income statement
                # if "income" in child_pat.lower() or "profit" in child_pat.lower():
                #     print("=" * 100)
                #     print("[DEBUG INCOME/P&L]")
                #     print("Child index:", idx)
                #     print("Child regex:", child_pat)
                #     print("Raw child pages:", child_pages)


                if r.get("report_type") == "AR":
                    co_pat = None
                    lower_pat = child_pat.lower()
                    if ("balance" in lower_pat and "sheet" in lower_pat) or ("financial" in lower_pat and "position" in lower_pat):
                        co_pat = SAME_PAGE_CO_TERMS_MAP_BS.get(idx)
                        # print(f"[BS] idx={idx}, co_pat={'SET' if co_pat else 'NONE'}") #debug
                        # print(f"[BS] child_pages={child_pages}") #debug
                    elif "cash" in lower_pat and "flow" in lower_pat:
                        co_pat = SAME_PAGE_CO_TERMS_MAP_CF.get(idx)
                    elif "income" in lower_pat or "profit" in lower_pat:
                        co_pat = SAME_PAGE_CO_TERMS_MAP_PL.get(idx)
                        # print(f"idx={idx}, co_pat={'SET' if co_pat else 'NONE'}") #debug

                    co_term_pages: List[int] = []  # fresh list per chil
                    if co_pat and child_pages:
                        try:
                            rx_co = compile_regex(co_pat)
                        except re.error as e:
                            print(f"[WARN] Invalid co-term regex for child {idx}: {co_pat}; {e}")
                            child_pages = []
                        else:
                            for p in child_pages:
                                page_text = (
                                    page_texts[p - 1] + "\n" + page_texts[p]
                                    if p < len(page_texts)
                                    else page_texts[p - 1]
                                )

                                # # ── CHANGED: income/profit looks back one page too ──────────────────
                                # if "income" in lower_pat or "profit" in lower_pat:
                                #     prev  = page_texts[p - 2] if p >= 2 else ""
                                #     curr  = page_texts[p - 1]
                                #     nxt   = page_texts[p] if p < len(page_texts) else ""
                                #     page_text = prev + "\n" + curr + "\n" + nxt

                                # # ── UNCHANGED: BS and CF only look forward ───────────────────────────
                                # else:
                                #     page_text = (
                                #         page_texts[p - 1] + "\n" + page_texts[p]
                                #         if p < len(page_texts)
                                #         else page_texts[p - 1]
                                #     )



                                normalized_keep_nl = _normalize_keep_newlines(page_text)
                                normalized_all = _normalize_all_whitespace(page_text)

                                # debug for Balance Sheet
                                # if ("balance" in lower_pat and "sheet" in lower_pat) or ("financial" in lower_pat and "position" in lower_pat):
                                #     print("-" * 100)
                                #     print(f"[BS] Checking page: {p}")
                                #     print(f"[BS] Page text sample:\n{normalized_all[:3000]}")
                                #     print(f"[BS] Co-term matched: {bool(rx_co.search(normalized_all))}")

                                # debug for profit loass and income statement
                                # if "income" in lower_pat or "profit" in lower_pat:
                                #     print("-" * 100)
                                #     print("Checking Income/P&L page:", p)
                                #     print("lower_pat:", lower_pat)          # ← add this
                                #     print("co_pat being used:", co_pat)     # ← add this
                                #     print("Page text sample:")
                                #     print(normalized_all[:3000])            # ← increase to 3000
                                #     print("Co-term matched:", bool(rx_co.search(normalized_all)))

                                co_pat1 = SAME_PAGE_CO_TERMS_MAP_BS.get(1)
                                has_all_three = (
                                    len(r["children"]) >= 3 and
                                    all(compile_regex(r["children"][i]).search(normalized_keep_nl) for i in range(3))
                                    # ↑ Balance Sheet, Income Statement AND Cash Flow all found on this same page
                                )
                                if has_all_three:
                                    # print(f"has_all_three: {has_all_three}")
                                    continue
                                if rx_co.search(normalized_all) and ("SEGMENT INFORMATION" not in normalized_all or "Independent Auditors’ Report" not in normalized_all) :
                                    print(f"[{p}] {idx}")
                                    co_term_pages.append(p)
                                    break
                            if "balance" in lower_pat and "sheet" in lower_pat:
                                raw_end, end_date = extract_ar_fiscal_end_date(page_texts)
                                if end_date:
                                    fiscalYearDate = "Date-" + raw_end
                                else:
                                    all_dt_set, all_dt_to_raw = extract_dates_from_text(normalized_all.replace("Notes ", "").replace("Notes ", ""))
                                    ctx_dt_set, ctx_dt_to_raw = extract_balance_sheet_context_dates(page_texts)
                                    if not all_dt_set:
                                        if ctx_dt_set:
                                            chosen = get_valid_max(ctx_dt_set)
                                            if chosen:
                                                raw = next(iter(ctx_dt_to_raw[chosen]))
                                                fiscalYearDate = "Date-" + raw
                                    elif len(all_dt_set) > 1:
                                        common = all_dt_set & ctx_dt_set if ctx_dt_set else set()
                                        if common:
                                            chosen = get_valid_max(common)
                                        else:
                                            chosen = get_valid_max(all_dt_set) or get_valid_max(ctx_dt_set)
                                        if chosen:
                                            raw = next(iter(ctx_dt_to_raw.get(chosen, all_dt_to_raw[chosen])))
                                            fiscalYearDate = "Date-" + raw
                                    else:
                                        chosen = next(iter(all_dt_set))
                                        if ctx_dt_set and chosen not in ctx_dt_set:
                                            chosen =get_valid_max(ctx_dt_set)
                                            raw = next(iter(ctx_dt_to_raw[chosen]))
                                        else:
                                            raw = next(iter(all_dt_to_raw[chosen]))
                                        fiscalYearDate = "Date-" + raw
                            child_pages = co_term_pages
                    if co_term_pages:
                        child_pages = co_term_pages
                    elif r.get("report_type") == "SR":
                        try:
                            rx_dateAR = compile_regex(dateAR) if dateAR else None
                            dates_pages = pages_for_regex(page_texts, rx_dateAR) if rx_dateAR else []
                            if dates_pages:
                                fiscalYearDate = "Date-" + dateAR
                            else:
                                ctx_dt_set, ctx_dt_to_raw = extract_balance_sheet_context_dates(page_texts)
                                if ctx_dt_set:
                                    chosen = get_valid_max(ctx_dt_set)
                                    raw = next(iter(ctx_dt_to_raw[chosen]))
                                    fiscalYearDate = raw + " - Not Matched with AR Report"
                        except re.error as e:
                            print(f"[WARN] Invalid dateAR regex '{dateAR}': {e}")
                children_pages_map[child_pat] = child_pages
                if child_pages:
                    matched_children.append(child_pat)

                # # debug
                # print(f"[POST] idx={idx}")
                # print(f"[POST] child_pages={child_pages}")
                # print(f"[POST] co_term_pages={co_term_pages}")
                # print(f"[POST] matched_children count={len(matched_children)}")
            # --- Satisfaction check ---
            children_matched = len(matched_children)

            #debug
            # print("Children required:", children_required)
            # print("Children matched:", children_matched)
            # print("Matched children:", matched_children)
            # print("Children pages map:", children_pages_map)
            # print("Status:", status)

            if require_all:
                satisfied = (children_matched == children_required)
                status = f"Parent matched; children {children_matched}/{children_required} (require ALL)."
            else:
                satisfied = (children_matched >= min_child_matches)
                status = f"Parent matched; children {children_matched} (require at least {min_child_matches})."
            if satisfied:
                return {
                    "is_correct": True,
                    "matched_row_index": r["row_index"],
                    "matched_parent": r["parent"],
                    "parent_pages": parent_pages,
                    "matched_children": matched_children,
                    "children_pages_map": children_pages_map,
                    "children_required": children_required,
                    "children_matched": children_matched,
                    "status": status,
                    "fiscalYearDate": fiscalYearDate,
                }
                break
    if "Parent found but after the page 10." in status1 :
        status = "Parent found but after the page 10."
    remark = ""
    if("No parent match for row" in status):
        return {
            "is_correct": False,
            "matched_row_index": None,
            "matched_parent": r["parent"],
            "parent_pages": parent_pages,
            "matched_children": matched_children,
            "children_pages_map": children_pages_map,
            "children_required": children_required,
            "children_matched": children_matched,
            "status": "No row satisfied for this report type",
            "fiscalYearDate": ""
        }
    elif("Parent found but after the page 10" in status):
        return {
            "is_correct": False,
            "matched_row_index": None,
            "matched_parent": r["parent"],
            "parent_pages": parent_pages,
            "matched_children": matched_children,
            "children_pages_map": children_pages_map,
            "children_required": children_required,
            "children_matched": children_matched,
            "status": "Parent found but after the page 10.",
            "fiscalYearDate": ""
        }
    else:
        image_pages = detect_image_pages(pdf_path)
        if len(image_pages)>=10:
            remark = "PDF Contains Images Please check"
        else :
            remark = "No row satisfied for this report type."
    # If nothing satisfied
    return {
        "is_correct": False,
        "matched_row_index": None,
        "matched_parent": r["parent"],
        "parent_pages": parent_pages,
        "matched_children": matched_children,
        "children_pages_map": children_pages_map,
        "children_required": children_required,
        "children_matched": children_matched,
        "status": remark,
        "fiscalYearDate": ""
    }

@lru_cache(maxsize=4096)
def compile_regex(pat: str):
    # NEW: cache regex compilation heavily used in matching
    flags = 0
    if GLOBAL_CASE_INSENSITIVE:
        flags = re.IGNORECASE
    flags |= re.MULTILINE  # make ^ and $ work per line
    return re.compile(pat, flags)

@lru_cache(maxsize=512)
def detect_image_pages(pdf_path):
    # NEW: cached image-page detection for speed on re-use
    doc = fitz.open(pdf_path)
    image_pages = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        has_image = any(block["type"] == 1 for block in blocks)  # type 1 = image block
        if has_image:
            image_pages.append(page_num + 1)
    return image_pages

def pages_for_regex(page_texts: list[str], rx):
    """Return sorted list of 1-based page numbers where rx matches."""
    return [i for i, t in enumerate(page_texts, start=1) if t and rx.search(t)]

def extract_balance_sheet_context_dates(page_texts):
    """
    Find pages with 'Balance Sheet as of/at' (or IFRS equivalent)
    and extract date candidates from the trailing segment.
    """
    # 1) Page locator: the phrase followed by any trailing characters
    phrase_rx = compile_regex(
        r"(?i)"  # case-insensitive
        r"(?:balance\s+sheet|"
        r"statement\s+of\s+financial\s+position|"
        r"consolidated\s+statement\s+of\s+financial\s+position|"
        r"Financial\s+Statement\s+for\s+the|"
        r"For\s+the\s+fiscal|"
        r"consolidated\s+financial\s+statements)\s+"
        r"(?:as\s+of|as\s+at|year\s+ended|year\s+ending)\s+"
        r".+"
    )
    bs_pages = pages_for_regex(page_texts, phrase_rx)
    # 2) Context extractor: capture the date-like segment after the phrase
    context_rx = re.compile(
        r"(?i)"  # case-insensitive
        r"(?:balance\s+sheet|"
        r"statement\s+of\s+financial\s+position|"
        r"consolidated\s+statement\s+of\s+financial\s+position|"
        r"Financial\s+Statement\s+for\s+the|"
        r"consolidated\s+financial\s+statements)\s+"
        r"(?:as\s+of|as\s+at|year\s+ended)\s+"
        r"((?:\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+,\s*\d{4})|"
        r"(?:\w{3,}\.?\s+\d{1,2},\s+\d{4})|"
        r"(?:\d{1,2}\s+[A-Za-z]+\s+\d{4})|"
        r"(?:[A-Za-z]+\s+\d{1,2},\s+\d{4})|"
        r"(?:\d{4}\.\d{2}\.\d{2})|"
        r"(?:\d{2}/\d{2}/\d{4}))"
    )
    ctx_dt_set = set()
    ctx_dt_to_raw = {}
    for bp in bs_pages:
        # Be defensive: if page text is missing, treat as empty string
        page_text_norm = _normalize_all_whitespace(page_texts[bp - 1] or "")
        m = context_rx.search(page_text_norm)
        if not m:
            continue
        # Use safe strip in case group(1) is ever None in other engines/refactors
        segment = sstrip(m.group(1))
        if not segment:
            continue
        seg_dt_set, seg_dt_to_raw = extract_dates_from_text(segment)
        ctx_dt_set |= seg_dt_set
        for k, v in seg_dt_to_raw.items():
            ctx_dt_to_raw.setdefault(k, set()).update(v)
    return ctx_dt_set, ctx_dt_to_raw

def _normalize_all_whitespace(s: str) -> str:
    # Collapse all whitespace (including newlines) to single spaces
    return re.sub(r"\s+", " ", s)

def sstrip(x) -> str:
    """
    Safe strip: returns '' for None and trims strings; otherwise str(x).strip().
    """
    if x is None:
        return ""
    return x.strip() if isinstance(x, str) else str(x).strip()

def extract_dates_from_text(text: str) -> List[str]:
    # Reusable month token: short and long names, optional dots for abbreviations, case-insensitive
    month_rx = r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|\
Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t\
tember)|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    date_rx = re.compile(rf"""
        \b(
            \d{{4}}\.\d{{2}}\.\d{{2}}  # 2024.12.31
            |{month_rx}\.? \s+ \d{{1,2}},? \s+ \d{{4}}  # Jan. 31, 2024 / January 31, 2024
            |{month_rx} \s+ \d{{4}}  # December 2023
            |\d{{2}}/\d{{2}}/\d{{4}}  # 31/01/2024 or 01/31/2024
            |\d{{1,2}} \s+ {month_rx} \s+ \d{{4}}  # 31 January 2024
            |\d{{1,2}}(?:st|nd|rd|th)? \s+ {month_rx} , \s* \d{{4}}  # 31st December, 2024
        )\b
    """, re.VERBOSE | re.IGNORECASE)
    candidates = {m.group(1).strip() for m in date_rx.finditer(text or "")}
    dt_set = set()
    dt_to_raw = {}
    for raw in candidates:
        if re.match(r"^\w{3,}\.?\s+\d{4}$", raw):
            try:
                parts = raw.replace('.', '').split()
                month = parts[0]
                year = int(parts[1])
                month_num = parser.parse(month).month
                last_day = monthrange(year, month_num)[1]
                raw = f"{last_day} {month} {year}"
            except Exception:
                pass
        if "and" not in raw:
            dt = try_parse_date(raw)
            if dt:
                dt_set.add(dt)
                dt_to_raw.setdefault(dt, set()).add(raw)
    dt_to_raw.setdefault(dt, set()).add(raw) if 'dt' in locals() and dt else None
    return dt_set, dt_to_raw

def try_parse_date(s: str):
    s_clean = re.sub(r'(\d{1,2})(st|nd|rd|th)', r'\1', s)
    for fmt in (
        "%d/%m/%Y",  # DD/MM/YYYY preferred
        "%m/%d/%Y",  # MM/DD/YYYY fallback
        "%b. %d, %Y",  # Jan. 31, 2024
        "%b %d, %Y",  # Jan 31, 2024
        "%d %B %Y",  # 31 January 2024
        "%B %d, %Y",  # January 31, 2024
        "%d %B, %Y",  # 31 December, 2024
        "%d %B,%Y",  # 31 December,2024
        "%Y.%m.%d",  # 2024.12.31
    ):
        try:
            return datetime.strptime(s_clean, fmt)
        except ValueError:
            continue
    # Optional fallback using dateutil
    try:
        return parser.parse(s_clean, dayfirst=True, fuzzy=False)
    except Exception:
        return None

def get_valid_max(date_set):
    MAX_YEAR = 2025
    valid_dates = [d for d in date_set if d.year <= MAX_YEAR]
    return max(valid_dates) if valid_dates else None

def extract_sr_fiscal_end_date(page_texts, max_pages=40):
    """
    Return (raw_phrase, dt_date) with improved priority order:
    1) 'About this Report' containing explicit date ranges (e.g., 'Oct 1, 2023 to Sep 30, 2024')
    2) 'About this Report' containing 'calendar year <YYYY>' -> Dec 31, <YYYY>
    3) 'period ... to <date>' or 'year ended <date>' / 'fiscal year ended <date>'
    4) conservative fallback (avoids footnote/criteria/baseline noise)
    """
    import re
    from datetime import date
    from dateutil import parser as dtparser
    # --- 1) Prefer explicit ranges on About page: '<date> to <date>' ---
    rx_about = re.compile(r"\bAbout\s+this\s+Report\b", re.IGNORECASE)
    # capture two date-like phrases separated by 'to'/'through'
    rx_range = re.compile(r"""(?ix)
        (?P<start>
        (?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
        jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|
        nov(?:ember)?|dec(?:ember)?) \s+ \d{1,2} \s*,?\s+ \d{4})
        \s* (?:to|through|–|-|—) \s*
        (?P<end>
        (?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
        jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|
        nov(?:ember)?|dec(?:ember)?) \s+ \d{1,2} \s*,?\s+ \d{4})
    """)
    for p in range(min(max_pages, len(page_texts))):
        text = page_texts[p] or ""
        if rx_about.search(text):
            m = rx_range.search(text)
            if m:
                end_raw = m.group("end")
                try:
                    end_dt = dtparser.parse(end_raw, dayfirst=False).date()
                    return f"{m.group(0)}", end_dt
                except Exception:
                    pass
    # --- 2) Calendar-year phrasing -> Dec 31 of that year ---
    rx_calendar_year = re.compile(r"(?:calendar\s+year)\s+(\d{4})", re.IGNORECASE)
    for p in range(min(max_pages, len(page_texts))):
        text = page_texts[p] or ""
        if rx_about.search(text):
            m = rx_calendar_year.search(text)
            if m:
                yy = int(m.group(1))
                return f"Calendar year {yy}", date(yy, 12, 31)
    # --- 3) Other common cues: 'year/fiscal year ended <date>' or '... to <date>' ---
    rx_year_ended = re.compile(r"""(?ix)
        (?:fiscal\s+year|year)\s+ended\s+
        ((?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
        jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|
        nov(?:ember)?|dec(?:ember)?)\s+\d{1,2}\s*,?\s+\d{4})
    """)
    rx_to_date = re.compile(r"""(?ix)
        (?:period|report(?:ing)?\s+period|boundary|scope).{0,60}?
        (?:to|through|ended)\s+
        ((?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
        jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|
        nov(?:ember)?|dec(?:ember)?)\s+\d{1,2}\s*,?\s+\d{4})
    """)
    for p in range(min(max_pages, len(page_texts))):
        text = page_texts[p] or ""
        for rx in (rx_year_ended, rx_to_date):
            m = rx.search(text)
            if m:
                raw = m.group(1)
                try:
                    return raw, dtparser.parse(raw, dayfirst=False).date()
                except Exception:
                    continue
    # --- 4) Conservative fallback near scope cues ---
    cue_rx = re.compile(r"(?i)\b(about\s+this\s+report|reporting\s+period|boundary|scope|fiscal\s+year|year\s+ended)\b")
    ban_rx = re.compile(r"(?i)\b(published|criteria|baseline|footnote|technical)\b")
    date_piece_rx = re.compile(r"""(?ix)\b(
        (?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|
        jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|
        nov(?:ember)?|dec(?:ember)?) \s+ \d{1,2} \s*,? \s+ \d{4}
    )\b""")
    for p in range(min(12, len(page_texts))):
        text = page_texts[p] or ""
        if cue_rx.search(text) and not ban_rx.search(text):
            m = date_piece_rx.search(text)
            if m:
                raw = m.group(1)
                try:
                    return raw, dtparser.parse(raw, dayfirst=False).date()
                except Exception:
                    pass
    return None, None

def _normalize_keep_newlines(s: str) -> str:
    # Collapse spaces/tabs but keep newlines intact
    return re.sub(r"[^\S\n]+", " ", s)

def extract_ar_fiscal_end_date(page_texts, max_pages=60):
    """
    Find 'For the fiscal year ended <date>' OR 'Fiscal year ended <date>'
    OR 'Year Ended <date>' within the first `max_pages` pages.
    Returns (raw_phrase, dt_date) or (None, None).
    """
    import re
    from dateutil import parser
    
    MONTHS_RX = r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|' \
            r'Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|' \
            r'Nov(?:ember)?|Dec(?:ember)?)'

    # Date component:  "Month DD, YYYY"  OR  "Month YYYY"  OR  "YYYY"
    DATE_RX = rf'(?:{MONTHS_RX}\s+\d{{1,2}},?\s*\d{{4}}|' \
          rf'{MONTHS_RX}\s+\d{{4}}|' \
          rf'\d{4})'

   # rx_for_fye = re.compile(r"(?i)for\s+the\s+fiscal\s+year\s+ended\s+(.{3,40})")
    #rx_fye     = re.compile(r"(?i)\bfiscal\s+year\s+ended\s+(.{3,40})")
    #rx_year_end= re.compile(r"(?i)\byear\s+ended\s+(.{3,40})")   # NEW
    
    rx_for_fye  = re.compile(rf"(?i)\bfor\s+the\s+fiscal\s+year\s+ended\s+(?P<date>{DATE_RX})(?=[,.;)]|\s|$)")
    rx_fye      = re.compile(rf"(?i)\bfiscal\s+year\s+ended\s+(?P<date>{DATE_RX})(?=[,.;)]|\s|$)")
    rx_year_end = re.compile(rf"(?i)\byear\s+ended\s+(?P<date>{DATE_RX})(?=[,.;)]|\s|$)")


    for p in range(min(max_pages, len(page_texts))):
        text = page_texts[p] or ""
        m = rx_for_fye.search(text) or rx_fye.search(text) or rx_year_end.search(text)
        if not m:
            continue

        raw = m.group(1).strip()
        # Trim after newline/section break
        #raw = re.split(r"[\n\r\\]", raw)[0].strip()
        # Remove ordinal suffixes (e.g., 31st)
        raw_clean = re.sub(r"\b(\d{1,2})(st|nd|rd|th)\b", r"\1", raw, flags=re.IGNORECASE)

        # Try robust parse (US-style month/day appears in 10-Ks)
        try:
            dt = parser.parse(raw_clean, dayfirst=False).date()
            return raw, dt
        except Exception:
            for fmt in ("%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%m/%d/%Y", "%Y-%m-%d", "%b %d, %Y"):
                try:
                    from datetime import datetime
                    dt = datetime.strptime(raw_clean, fmt).date()
                    return raw, dt
                except Exception:
                    pass

    return None, None

def normalize_pdf_text(s: str) -> str:
    # Unicode normalization (canonical/compatibility)
    s = unicodedata.normalize('NFKC', s)
    # Replace special spaces & dashes, remove zero-width chars
    s = s.translate(TRANS)
    # Normalize newlines and collapse extra whitespace
    s = s.replace('\r', '\n')
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r'\n{2,}', '\n', s)
    return s.strip()

def format_children_pages(children_pages_map: dict) -> str:
    """Format as 'label: 1,3; another: 2'. (FIXED: default label + typo)"""
    if not children_pages_map:
        return ""
    parts = []
    for pat in children_pages_map.keys():
        lower_pat = pat.lower()
        # Default to the raw pattern if no known label matched
        pat1 = pat
        if ("balance" in lower_pat and "sheet" in lower_pat) or ("financial" in lower_pat and "position" in lower_pat):
            pat1 = "Balance Sheet Table"
        elif "cash" in lower_pat and "flow" in lower_pat:
            pat1 = "Cash Flow Table"
        elif ("income" in lower_pat) or ("profit" in lower_pat):
            pat1 = "Income statement"  # typo fixed
        pages = children_pages_map.get(pat, [])
        pages_s = ",".join(map(str, pages)) if pages else ""
        parts.append(f"{pat1}: {pages_s}")
    return "; ".join(parts)

def convert_to_ddmmyyyy(x):
    # 1) Missing checks (covers NaT, NaN, None, <NA>)
    if pd.isna(x):
        return pd.NA
    # 2) Datetime-like objects
    if isinstance(x, (pd.Timestamp, datetime)):
        return x.strftime("%d%m%Y")
    if hasattr(x, "to_pydatetime"):
        try:
            return x.to_pydatetime().strftime("%d%m%Y")
        except Exception:
            pass
    # 3) Numeric Excel serial dates
    if isinstance(x, (int, float)) and not np.isnan(x):
        try:
            dt = pd.to_datetime(x, origin="1899-12-30", unit="D")  # Excel origin
            return dt.strftime("%d%m%Y")
        except Exception:
            pass
    # 4) Strings (clean and parse)
    s = str(x).strip()
    if s == "" or s.lower() in {"nat", "nan", "none", "null"}:
        return pd.NA
    if s.lower().startswith("date-"):
        s = s[5:].strip()
    if "- Not Matched with AR Report".lower() in s.lower():
        s = s.split(" - ")[0]
    # Normalize ordinal suffixes like "31st March 2024"
    s = pd.Series([s]).str.replace(r"\b(\d{1,2})(st|nd|rd|th)\b", r"\1", regex=True).iloc[0]
    # If it's plain 8-digit DDMMYYYY, keep as is (validate)
    if pd.Series([s]).str.match(r"^\d{8}$").iloc[0]:
        try:
            dt = datetime.strptime(s, "%d%m%Y")
            return dt.strftime("%d%m%Y")
        except ValueError:
            return pd.NA
    for fmt in (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d %B %Y",  # 31 March 2024
        "%b %d, %Y",  # Mar 31, 2024
        "%B %d, %Y",  # March 31, 2024
        "%d %b %Y",  # 31 Mar 2024
    ):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d%m%Y")
        except ValueError:
            continue
    # Last attempt: let pandas parse flexibly (dayfirst=True)
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return dt.strftime("%d%m%Y")
    except Exception:
        return pd.NA  # Not parseable

def validate_year(date_str, threshold=2024):
    """
    Returns (is_valid, message_or_None).
    - If year < threshold -> (False, "Year is less than 2024")
    - If year >= threshold and date between Jan-Jun 2024 -> (False, "This is last year report.")
    - If year >= threshold and not in Jan-Jun 2024 -> (True, None)
    - If not parseable -> (False, "Invalid or missing year")
    """
    try:
        if pd.isna(date_str):
            return False, None
        s = str(date_str).strip()
        if s.isdigit() and len(s) == 8:
            s = f"{s[0:2]}-{s[2:4]}-{s[4:]}"  # Convert to DD-MM-YYYY
        # Second check: date between Jan-Jun 2024
        parsed_date = pd.to_datetime(s, dayfirst=True, errors='raise')
        if datetime(2024, 1, 1) <= parsed_date <= datetime(2024, 6, 30):
            return False, "This is last year report."
        # Try parsing date
        try:
            parsed_date = pd.to_datetime(s, errors='raise')
        except Exception:
            # If parsing fails, fallback to year check
            if len(s) < 4 or not s[-4:].isdigit():
                return False, "Invalid or missing year"
            year = int(s[-4:])
            if year < threshold:
                return False, "Year is less than 2024"
            else:
                return True, None
        # Extract year
        year = parsed_date.year
        # First check: year < threshold
        if year < threshold:
            return False, "Year is less than 2024"
        # Otherwise valid
        return True, None
    except Exception:
        return False, "Invalid or missing year"

def generate_summary(dfPortfolioTracker):
    summary = []
    # Total PDFs
    total_pdfs = len(dfPortfolioTracker)
    summary.append(["Total PDFs", "", total_pdfs])
    # pdf_report_not_found (Remarks contains "PDF not found")
    pdf_report_not_found = dfPortfolioTracker['Remarks'].str.contains("PDF not found", case=False, na=False).sum()
    summary.append(["PDF Report not found ", "", pdf_report_not_found])
    # Correct PDFs by Validation == TRUE (PATCHED)
    correct_pdfs = (dfPortfolioTracker['Validation'].astype(str).str.upper() == "TRUE").sum()
    summary.append(["Correct PDFs", "", correct_pdfs])
    # Prepare columns
    remarks = dfPortfolioTracker['Remarks'].fillna("").astype(str)
    # Conditions for sub-metrics
    conditions = {
        "Company Name Matched": remarks.str.contains("Parent & Company Name Matched", case=False, na=False),
        "Company Name Missmatched": remarks.str.contains("Parent matched but Company name Mismatch", case=False, na=False),
        "Modified Date > 2024": remarks.str.contains("Modified year is > 2024", case=False, na=False),
        "Created Date < 2025": remarks.str.contains("Created Date is < 2025", case=False, na=False),
    }
    # Count each condition
    for sub_metric, condition in conditions.items():
        count = condition.sum()
        summary.append(["Correct PDFs", sub_metric, count])
    # Incorrect PDFs
    incorrect_pdfs = total_pdfs - pdf_report_not_found - correct_pdfs
    summary.append(["Incorrect PDFs", "", incorrect_pdfs])
    # Conditions for sub-metrics
    conditions = {
        "Corrupted PDF": remarks.str.contains("Corrupted PDF - Unable to open", case=False, na=False),
        "Year is less than 2024": remarks.str.contains("Year is less than 2024", case=False, na=False),
        "Fiscal Year Date Not found": remarks.str.contains(r"(Fiscal Year Date not found|Parent matched - Only Fiscal Year found in report)",case=False, na=False),
        "Reports not matched with the Key Words": remarks.str.contains("No row satisfied for this report type", case=False, na=False) \
            | remarks.str.contains("Proper Financial Statement key word not available within the 3 pages", case=False, na=False)\
            | remarks.str.contains("Quarterly report detected", case=False, na=False),
        "Image PDF": remarks.str.contains("PDF Contains Images Please check", case=False, na=False),
        "Parent Key Word after Page no. 10": remarks.str.contains("Parent found but after the page 10.", case=False, na=False),
        "Non-English Report": remarks.str.contains("PDF Language is", case=False, na=False),
        "Small PDF": remarks.str.contains("PDF pages is less than 10", case=False, na=False),
        "Last year Report": (remarks.str.contains("This is last year report.", case=False, na=False)& ~remarks.str.contains("Modified year is > 2024", case=False, na=False)),
    }
    for sub_metric, condition in conditions.items():
        count = condition.sum()
        summary.append(["Incorrect PDFs", sub_metric, count])
    matched_any = pd.concat(conditions.values(), axis=1).any(axis=1)
    unknown_type_count = ((dfPortfolioTracker['Validation'].astype(str).str.upper() == "FALSE") & (~matched_any)).sum()
    summary.append(["Incorrect PDFs", "Unknown_type", unknown_type_count])
    report_summary = pd.DataFrame(summary, columns=["Metric", "Sub-Metric", "Count"])
    return report_summary

if __name__ == "__main__":
    # folder = r"C:\Users\janarddanm\Downloads\1_Validation Analysis\google_Search_reports - Wrong False\Allianz SE_AR"
    folder = r"C:\Users\janarddanm\Downloads\1_Validation Analysis\google_Search_reports all"
    results, output_path = worker(folder)
    print("Output file:", output_path)
    print("Total processed:", len(results))




def validate_pdf(file_path: str, entity_name: str, report_type: str, km_rows_all: list,report_name: str, c4f_entity_id: str) -> dict:
    """
    Returns:
        {
            "is_valid":                    True/False,
            "Validation":                  "Correct" / "Incorrect" / "Error" etc,
            "Remarks":                     "reason string",
            "Language":                    "en" / "fr" / "No text found" etc,
            "Year":                        "2024" etc or "",
            "Fiscal Year Date (DDMMYYYY)": "31032024" etc or "",
            "Created Date (DDMMYYYY)":     "01012024" etc or "",
            "Modified Date (DDMMYYYY)":    "01012024" etc or "",
        }
    """

    print("\n[validate_pdf] Inputs:")
    print("-" * 65)
    print(f"{'Field':<20} | {'Value'}")
    print("-" * 65)
    print(f"{'file_path':<20} | {file_path}")
    print(f"{'entity_name':<20} | {entity_name}")
    print(f"{'report_type':<20} | {report_type}")
    print(f"{'km_rows_all':<20} | {len(km_rows_all)} rows")
    print(f"{'report_name':<20} | {report_name}")
    print(f"{'c4f_entity_id':<20} | {c4f_entity_id}")
    print("-" * 65)

    # ── Default return structure ──────────────────────────────────────────
    default = {
        "is_valid":                    False,
        "Validation":                  "",
        "Remarks":                     "",
        "Language":                    "",
        "Year":                        "NA",
        "Fiscal Year Date (DDMMYYYY)": "",
        "Created Date (DDMMYYYY)":     "",
        "Modified Date (DDMMYYYY)":    "",
        "Report Name":                 report_name,
        "C4F Entity ID":               c4f_entity_id,
    }

    name          = os.path.basename(file_path)
    language      = ""
    created_date  = "Date not available"
    modified_date = "Date not available"
    inferred_type = None
    page_texts    = []
    res = {
        "fiscalYearDate":    "",
        "is_correct":        False,
        "matched_row_index": None,
        "matched_parent":    "",
        "parent_pages":      [],
        "matched_children":  [],
        "children_pages_map":{},
        "children_required": "",
        "children_matched":  "",
        "status":            ""
    }

    # ── Step 1: Metadata + language checks ───────────────────────────────
    try:
        inferred_type = parse_report_type_from_name(name)

        reader = PdfReader(file_path)

        if getattr(reader, "is_encrypted", False):
            try:
                reader.decrypt("")
            except Exception:
                pass

        metadata      = reader.metadata or {}
        created_date  = convert_pdf_date(metadata.get("/CreationDate"))
        modified_date = convert_pdf_date(metadata.get("/ModDate"))

        # ── Page count check ─────────────────────────────────────────────
        if len(reader.pages) < 10:
            return {
                **default,
                "Remarks":                     "PDF pages is less than 10",
                "Created Date (DDMMYYYY)":     created_date,
                "Modified Date (DDMMYYYY)":    modified_date,
            }

        # ── Language check ───────────────────────────────────────────────
        text, has_text = fast_extract_pdf_text(file_path)

        if has_text:
            try:
                language = detect(text)
            except LangDetectException:
                language = ""
                return {
                    **default,
                    "Remarks":                  "Language detection failed",
                    "Language":                 "",
                    "Created Date (DDMMYYYY)":  created_date,
                    "Modified Date (DDMMYYYY)": modified_date,
                }

            if language != "en":
                return {
                    **default,
                    "Remarks":                  "Non-English Report",
                    "Language":                 language,
                    "Created Date (DDMMYYYY)":  created_date,
                    "Modified Date (DDMMYYYY)": modified_date,
                }
        else:
            language = "No text found"
            return {
                **default,
                "Remarks":                  "PDF Contains Images Please check",
                "Language":                 language,
                "Created Date (DDMMYYYY)":  created_date,
                "Modified Date (DDMMYYYY)": modified_date,
            }

    except PdfReadError as e:
        return {
            **default,
            "Remarks":                  "Corrupted PDF - Unable to open",
            "Language":                 "Language detection skipped due to corrupted PDF",
            "Created Date (DDMMYYYY)":  "Date not available",
            "Modified Date (DDMMYYYY)": "Date not available",
        }

    except Exception as e:
        language = f"Language detection skipped: {e}"

    # ── Step 2: Main PDF matching logic ──────────────────────────────────
    try:
        page_texts = extract_text_per_page(file_path)

        # ── Wrong report type check ──────────────────────────────────────
        if inferred_type in ("AR", "SR"):
            first_page_text = page_texts[0] if page_texts else ""
            if first_page_text == "":
                first_page_text = page_texts[1] if len(page_texts) > 1 else ""

            is_wrong_type = re.search(
                r"(Quarterly\s+Report|Interim\s+Report|Half[\s-]?Year"
                r"|\b10-?Q\b|\b6-?K\b|\b8-?K\b"
                r"|Immediate\s+Report|Proxy\s+Statement|Prospectus"
                r"|\bSupplement(?:al|ary)?\b(?!\s+(?:ing\s+the\s+)?(?:Annual|Sustainability|ESG|Integrated)\s+Report)"
                r"|(Annual|Sustainability|ESG|Integrated|Corporate)\s+Report\s+(19[0-9]{2}|20[01][0-9]|202[0-4])\b"
                r"|\bFY\s?(19[0-9]{2}|20[01][0-9]|202[0-4])\b"
                r"|(19[0-9]{2}|20[01][0-9]|202[0-4])\s+(Annual|Sustainability|ESG|Corporate|ENVIRONMENTAL)\b)",
                first_page_text,
                re.IGNORECASE
            )

            if is_wrong_type:
                return {
                    **default,
                    "Remarks":                  "Incorrect or Old report",
                    "Language":                 language,
                    "Created Date (DDMMYYYY)":  created_date,
                    "Modified Date (DDMMYYYY)": modified_date,
                }

        # ── KeyMaster matching ───────────────────────────────────────────
        if inferred_type is None:
            res["status"] = "Could not infer report type from filename"

        else:
            rows_subset = [r for r in km_rows_all if r["report_type"] == inferred_type]

            if not rows_subset:
                res["status"] = f"No KeyMaster rows for report type '{inferred_type}'."

            else:
                # ── SR: get AR fiscal year date from validate Parquet ────
                # In old code this came from dfPortfolioTracker
                # Pass dateAR from outside if needed, default empty for now
                dateAR = ""

                if res["status"] == "" and language == "en":
                    res = match_pdf_with_rows(
                        page_texts,
                        rows_subset,
                        dateAR,
                        file_path
                    )

        record = build_record(
            name,
            file_path,
            inferred_type,
            res,
            created_date,
            modified_date
        )

    except Exception as e:
        return {
            **default,
            "Remarks":                  f"[ERROR] {e}",
            "Language":                 language,
            "Created Date (DDMMYYYY)":  created_date,
            "Modified Date (DDMMYYYY)": modified_date,
        }

    # ── Step 3: Fiscal year date parsing ─────────────────────────────────
    date_obj       = None
    fiscalYearDate1 = ""
    fy_val         = str(record.get("fiscalYearDate", "") or "")
    stat           = record.get("is_correct")
    year           = "NA"

    if fy_val and "Not Matched with AR Report" not in fy_val and "Date-" in fy_val:
        raw_date = fy_val.replace("Date-", "").strip()
        s = str(raw_date).strip()

        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%b. %d, %Y", "%b %d, %Y",
                    "%d %B %Y", "%B %d, %Y", "%d%B%Y"):
            try:
                date_obj = datetime.strptime(s, fmt)
                break
            except ValueError:
                continue

        if date_obj is None:
            normalized = "".join(ch for ch in s if ch.isalnum())
            for fmt in ("%d%m%Y", "%m%d%Y", "%Y%m%d"):
                try:
                    date_obj = datetime.strptime(normalized, fmt)
                    break
                except ValueError:
                    continue

        if date_obj:
            fiscalYearDate1 = date_obj.strftime("%d%m%Y")
            if len(fiscalYearDate1) == 8 and fiscalYearDate1.isdigit():
                year = int(fiscalYearDate1[-4:])

    elif re.match(r"^\d{4}$", str(fy_val)):
        year = fy_val

    # ── Step 4: Company name check ────────────────────────────────────────
    if stat is True:
        company_name = " ".join(
            entity_name.lower()
            .replace("_ar", "").replace("_sr", "").replace(".pdf", "")
            .replace("corp.", "").replace("corp", "").replace("ltd.", "")
            .replace("the ", "").replace("inc.", "").replace(" ", "")
            .replace(",", "").replace(".", "").replace("&", "").replace("'", "")
            .split()[:2]
        )

        pages_to_search = page_texts[:10] #raw_first_10 page
        if len(page_texts) > 10:
            pages_to_search = pages_to_search + page_texts[-2:] #add last 2 page text also

        raw_first_10 = " ".join(pages_to_search)

        first_10_pages_text = (
            normalize_pdf_text(raw_first_10)
            .lower()
            .replace("_ar", "").replace("_sr", "").replace(".pdf", "")
            .replace("corp.", "").replace("corp", "").replace("ltd.", "")
            .replace("the ", "").replace("inc.", "").replace(" ", "")
            .replace(",", "").replace(".", "").replace("&", "").replace("'", "")
            .replace("\n", "")
        )

        if company_name and company_name.lower() not in first_10_pages_text.lower():
            validation_status = "Parent matched but Company name Mismatch"
        else:
            validation_status = "Parent & Company Name Matched"
    else:
        validation_status = res.get("status", "")


        # ── Step 6: Copy and rename validated PDF if correct ─────────────────
    if stat is True:
        try:
            pdf_folder       = os.path.dirname(file_path)          # same folder as the PDF
            corrected_folder = os.path.join(pdf_folder, "Corrected Reports")
            os.makedirs(corrected_folder, exist_ok=True)

            actual_name = (
                c4f_entity_id.strip()
                + "_"
                + str(year).strip()
                + "_"
                + report_type.strip()
                + "_"
                + entity_name.strip()
                .replace("_AR", "").replace("_SR", "")
                .replace(".pdf", "")
            )

            target_path = os.path.join(corrected_folder, f"{actual_name}.pdf")

            if os.path.isfile(file_path):
                shutil.copy2(file_path, target_path)
                print(f"[Validate] Copied corrected PDF → {actual_name}.pdf")

        except Exception as e:
            print(f"[Validate] Failed to copy corrected PDF for {entity_name}: {e}")

    # ── Step 5: Build final return dict ──────────────────────────────────
    return {
        "is_valid":                    stat is True,
        "Validation":                  "Correct" if stat is True else "Incorrect",
        "Remarks":                     validation_status,
        "Language":                    language,
        "Year":                        str(year),
        "Fiscal Year Date (DDMMYYYY)": fiscalYearDate1,
        "Created Date (DDMMYYYY)":     created_date,
        "Modified Date (DDMMYYYY)":    modified_date,
        "Report Name":                 report_name,
        "C4F Entity ID":               c4f_entity_id,
    }
